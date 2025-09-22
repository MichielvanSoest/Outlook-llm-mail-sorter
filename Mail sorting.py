"""
This script connects to Outlook, grabs emails from a source folder,
and asks a local LLM (via LM Studio) where to put them.
It can create new folders if allowed, or dump into "Unsorted".
A log file is kept of everything that happens, so you know what went where.
Progress is shown with a progress bar, because waiting in silence sucks and Yes, I talk to a bot to sort my emails because humans are unreliable
"""

import re
import time
import unicodedata

import requests
from tqdm import tqdm
import win32com.client
from typing import Optional

from Secrets import (
    MODEL_ID, LMSTUDIO_URL, NUM_EMAILS, CREATE_NEW_FOLDERS,
    SOURCE_FOLDER_PATH, LOG_FILE_PATH, PERSONAL_DETAILS, EXAMPLE_PROMPT
)

# Constants
MAIL_ITEM_CLASS = 43
FORBIDDEN_PATTERNS = [
    r"postvak in[^/]*[^/]",   # e.g. "Postvak In Something" without slash
    r"\.\.\.",                # prevent weird "..." names
    r"[^a-zA-Z0-9/ \-]",      # block strange symbols (keep it clean)
    r"postvak in opr",        # specific bad case
]

# Cache of known folders (so we don’t have to re-scan)
known_folders = {}


# ---------- Folder helpers ----------

def normalize_name(name: str) -> str:
    # Strip accents, lowercase, clean whitespace
    # AKA "make it boring enough that computers won’t complain"
    nfkd = unicodedata.normalize("NFKD", name)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).strip().lower()


def build_folder_cache(root, base="Postvak In"):
    # Walk through Outlook folder tree and remember them
    for f in root.Folders:
        path = f"{base}/{f.Name}"
        known_folders[normalize_name(path)] = path
        build_folder_cache(f, path)


def get_or_create_folder(root_folder, folder_path: str):
    # Return folder object (create it if missing and allowed)
    norm = normalize_name(folder_path)
    if norm in known_folders:
        folder_path = known_folders[norm]

    parts = folder_path.split("/")
    folder = root_folder
    real_path = parts[0] if parts else "Postvak In"

    for part in parts[1:]:
        subfolder = next((f for f in folder.Folders if f.Name == part), None)
        if not subfolder:
            if CREATE_NEW_FOLDERS:
                print(f"Creating folder: {part} under {real_path}")
                subfolder = folder.Folders.Add(part)
            else:
                print(f"[INFO] Not creating '{part}', CREATE_NEW_FOLDERS=False. Sending to Unsorted.")
                return next((f for f in root_folder.Folders if f.Name == "Unsorted"), root_folder)
        else:
            print(f"Found existing folder: {part} under {real_path}")
        folder = subfolder
        real_path += "/" + part
        known_folders[normalize_name(real_path)] = real_path

    print(f"Final target folder: {real_path}")
    return folder

def extract_useful_metadata(raw_metadata):
    useful = []
    if "attachment_names" in raw_metadata:
        useful.append(f"Bijlagen: {', '.join(raw_metadata['attachment_names'])}")
    if "from_domain" in raw_metadata:
        useful.append(f"Domein afzender: {raw_metadata['from_domain']}")
    if "labels" in raw_metadata:
        useful.append(f"Labels: {', '.join(raw_metadata['labels'])}")
    return "\n".join(useful)

# ---------- Classification ----------
 # If the model fails or is lazy, fallback to Unsorted, but only if creating folders is not allowed to suggest better names

def classify_email(subject: str, body: str, sender_name: str,
                   sender_email: str, to: str, cc: str, attachments: list,
                   received_date: Optional[str] = None) -> str:
    """
    Classify an email into the correct folder using LLM.
    """

    attachments_info = ", ".join(attachments) if attachments else "None"
    existing_folders_list = "\n".join(sorted(set(known_folders.values())))

    # Optional metadata
    metadata = []
    if received_date:
        metadata.append(f"Ontvangen datum: {received_date}")
    metadata.append(f"Aantal bijlagen: {len(attachments)}")
    if attachments:
        metadata.append(f"Bijlagen namen/types: {attachments_info}")

    metadata_info = "\n".join(metadata)

    prompt = (
        f"Je bent een e-mail sorteerassistent. {PERSONAL_DETAILS}\n"
        f"{EXAMPLE_PROMPT.format(existing_folders_list=existing_folders_list)}"
        f"\nE-mail details:\n"
        f"Onderwerp: {subject}\n"
        f"Afzender: {sender_name} <{sender_email}>\n"
        f"Aan: {to}\n"
        f"CC: {cc}\n"
        f"{extract_useful_metadata(metadata)}\n"
        f"Inhoud: {body[:1000]}\n"
        f"\nFolder:\nGeef alleen het folderpad terug, zonder extra uitleg.\n"
    )

    response = requests.post(LMSTUDIO_URL, json={
        "model": MODEL_ID,
        "prompt": prompt,
        "max_tokens": 50
    })

    if response.status_code == 200:
        result = response.json()
        folder = result["choices"][0]["text"].strip()
        return folder if folder else "Postvak In/Unsorted"
    else:
        raise RuntimeError(f"LLM API error {response.status_code}: {response.text}")



# ---------- Mail handling ----------

def load_messages(source_folder, log_file):
    # Load mails, skip invalid ones
    raw_messages = [msg for msg in source_folder.Items if msg.Class == MAIL_ITEM_CLASS]
    messages, skipped = [], 0
    for msg in raw_messages:
        try:
            _ = msg.ReceivedTime
            messages.append(msg)
        except Exception:
            skipped += 1
            subj = getattr(msg, 'Subject', 'No Subject')
            print(f"[SKIP] No ReceivedTime: {subj}")
            log_file.write(f"SKIPPED: '{subj}' -> No ReceivedTime\n")
    return messages, skipped


def clean_folder_path(folder_path: str) -> str:
    # Normalize and sanitize folder name
    clean_path = folder_path.strip()
    clean_path = re.sub(r"\s+", " ", clean_path)

    if not clean_path.lower().startswith("postvak in/"):
        clean_path = f"Postvak In/{clean_path}"

    # Block broken names
    if any(re.search(p, clean_path.lower()) for p in FORBIDDEN_PATTERNS):
        print(f"[ERROR] Suggested folder '{folder_path}' invalid. Using 'Postvak In/Unsorted'.")
        clean_path = "Postvak In/Unsorted"

    return clean_path


def process_message(mail, source_folder, log_file):
    # Process a single mail (classify, move, log)
    subject = mail.Subject or ""
    body = mail.Body or ""
    sender_name = mail.SenderName or ""
    sender_email = mail.SenderEmailAddress or ""
    to = mail.To or ""
    cc = mail.CC or ""
    attachments = [att.FileName for att in mail.Attachments]

    print(f"\nProcessing: {subject}")

    folder_path = classify_email(
        subject=subject, body=body,
        sender_name=sender_name, sender_email=sender_email,
        to=to, cc=cc, attachments=attachments,
        received_date=str(mail.ReceivedTime)  # YYYY-MM-DD will work fine
    )

    print(f" -> Suggested folder: {folder_path}")

    folder_path = clean_folder_path(folder_path)

    if not CREATE_NEW_FOLDERS:
        folder_path = f"{SOURCE_FOLDER_PATH}/Unsorted"

    target_folder = get_or_create_folder(source_folder, folder_path)

    # Only move to mail-type folders... no, you can’t move it into the Contacts folder which somehow got suggested by the llm. must be my prompting abili"
    if hasattr(target_folder, 'DefaultItemType') and target_folder.DefaultItemType == 0:
        print(f"Moving mail to folder: {target_folder.Name}")
        mail.Move(target_folder)
        log_file.write(f"SORTED: '{subject}' -> '{folder_path}'\n")
    else:
        print(f"Target '{target_folder.Name}' is not a Mail folder. Skipped.")
        log_file.write(f"SKIPPED: '{subject}' -> '{folder_path}' (Not a Mail folder)\n")


# ---------- Main ----------

def main():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Helper to resolve "path style" folders
    def get_folder_by_path(root, path):
        folder = root
        for part in path.split("/"):
            folder = next((f for f in folder.Folders if f.Name == part), folder)
        return folder

    root_folder = outlook.Folders.Item(1)  # Default mailbox
    source_folder = get_folder_by_path(root_folder, SOURCE_FOLDER_PATH)

    print(f"Accessing folder: {SOURCE_FOLDER_PATH}")
    print(f"Resolved folder: {getattr(source_folder, 'Name', str(source_folder))}")

    build_folder_cache(source_folder)

    with open(LOG_FILE_PATH, "a", encoding="utf-8") as log_file:
        messages, skipped = load_messages(source_folder, log_file)

        print(f"Found {len(messages)} mails (skipped {skipped})")
        log_file.write(f"SUMMARY: {len(messages)} mails, {skipped} skipped\n")

        messages.sort(key=lambda x: x.ReceivedTime, reverse=True)
        messages = messages[:NUM_EMAILS]

        pbar = tqdm(total=len(messages), desc="Sorting emails", unit="email")
        start_time = time.time()

        for idx, mail in enumerate(messages, 1):
            if not hasattr(mail, 'Subject'):
                log_line = "SKIPPED: <unknown item> -> No Subject"
                print(log_line)
                log_file.write(log_line + "\n")
                continue

            try:
                process_message(mail, source_folder, log_file)
            except Exception as e:
                subj = getattr(mail, 'Subject', 'No Subject')
                print(f"Error: {e}")
                log_file.write(f"SKIPPED: '{subj}' -> ERROR: {e}\n")

            # Progress bar handles ETA itself, but we add elapsed time for fun
            elapsed = time.time() - start_time
            remaining = (elapsed / idx) * (len(messages) - idx) if idx > 0 else 0
            pbar.set_postfix({"Elapsed": f"{elapsed:.1f}s", "ETA": f"{remaining:.1f}s"})
            pbar.update(1)

        pbar.close()


if __name__ == "__main__":
    main()
