import win32com.client
import requests
import unicodedata
from tqdm import tqdm
import time
from Secrets import MODEL_ID, LMSTUDIO_URL, NUM_EMAILS, CREATE_NEW_FOLDERS, SOURCE_FOLDER_PATH, LOG_FILE_PATH, PERSONAL_DETAILS, EXAMPLE_PROMPT

# Cache of known folders (normalized -> real path)
known_folders = {}

def normalize_name(name: str) -> str:
    nfkd = unicodedata.normalize("NFKD", name)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).strip().lower()

def build_folder_cache(root, base="Postvak In"):
    for f in root.Folders:
        path = f"{base}/{f.Name}"
        known_folders[normalize_name(path)] = path
        build_folder_cache(f, path)

def classify_email(subject: str, body: str, sender_name: str, sender_email: str, to: str, cc: str, attachments: list) -> str:
    attachments_info = ", ".join(attachments) if attachments else "None"
    # List existing folders for the LLM
    existing_folders_list = "\n".join(sorted(set(known_folders.values())))
    prompt = (
        f"Je bent een e-mail sorteerassistent. {PERSONAL_DETAILS}\n"
        f"{EXAMPLE_PROMPT.format(existing_folders_list=existing_folders_list)}"
        f"\nE-mail details:\n"
        f"Onderwerp: {subject}\n"
        f"Afzender: {sender_name} <{sender_email}>\n"
        f"Aan: {to}\n"
        f"CC: {cc}\n"
        f"Bijlagen: {attachments_info}\n"
        f"Inhoud: {body[:1000]}\n"
        f"\nFolder:\nGeef alleen het folderpad terug, zonder extra uitleg.\n"
    )

    data = {
        "model": MODEL_ID,
        "prompt": prompt,
        "max_tokens": 50
    }

    response = requests.post(LMSTUDIO_URL, json=data)
    if response.status_code == 200:
        result = response.json()
        folder = result["choices"][0]["text"].strip()
        return folder if folder else "Postvak In/Unsorted"
    else:
        raise RuntimeError(f"LLM API error {response.status_code}: {response.text}")

def get_or_create_folder(root_folder, folder_path: str):
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
                print(f"[INFO] CREATE_NEW_FOLDERS is False. Not creating folder '{part}'. Returning 'Unsorted'.")
                return next((f for f in root_folder.Folders if f.Name == "Unsorted"), root_folder)
        else:
            print(f"Found existing folder: {part} under {real_path}")
        folder = subfolder
        real_path += "/" + part
        known_folders[normalize_name(real_path)] = real_path

    print(f"Final target folder: {real_path}")
    return folder

def main():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # Find source folder by path
    def get_folder_by_path(root, path):
        parts = path.split("/")
        folder = root
        for part in parts:
            folder = next((f for f in folder.Folders if f.Name == part), folder)
        return folder

    root_folder = outlook.Folders.Item(1)  # Usually the default account
    source_folder = get_folder_by_path(root_folder, SOURCE_FOLDER_PATH)

    print(f"Accessing folder: {SOURCE_FOLDER_PATH}")
    print(f"Resolved folder name: {getattr(source_folder, 'Name', str(source_folder))}")

    build_folder_cache(source_folder)

    # Open log file for writing
    log_file = open(LOG_FILE_PATH, "a", encoding="utf-8")
    raw_messages = [msg for msg in source_folder.Items if msg.Class == 43]  # 43 = MailItem
    messages = []
    skipped = 0
    for msg in raw_messages:
        try:
            _ = msg.ReceivedTime
            messages.append(msg)
        except Exception:
            print(f"[SKIP] Message without valid ReceivedTime: {getattr(msg, 'Subject', 'No Subject')}")
            skipped += 1
            log_file.write(f"SKIPPED: '{getattr(msg, 'Subject', 'No Subject')}' -> No ReceivedTime\n")
    print(f"Found {len(messages)} mail(s) in folder '{SOURCE_FOLDER_PATH}' (skipped {skipped} zonder ReceivedTime)")
    log_file.write(f"SUMMARY: Found {len(messages)} mail(s) in folder '{SOURCE_FOLDER_PATH}' (skipped {skipped} zonder ReceivedTime)\n")
    messages.sort(key=lambda x: x.ReceivedTime, reverse=True)
    messages = messages[:NUM_EMAILS]

    # Setup progress bar only AFTER messages is defined
    total_mails = len(messages)
    pbar = tqdm(total=total_mails, desc="Sorting emails", unit="email")
    start_time = time.time()
    for idx, mail in enumerate(messages, 1):
        # Skip items without a Subject attribute
        if not hasattr(mail, 'Subject'):
            log_line = f"SKIPPED: <unknown item> -> No Subject attribute"
            print(log_line)
            log_file.write(log_line + "\n")
            continue
        try:
            subject = mail.Subject or ""
            body = mail.Body or ""
            sender_name = mail.SenderName or ""
            sender_email = mail.SenderEmailAddress or ""
            to = mail.To or ""
            cc = mail.CC or ""
            attachments = [att.FileName for att in mail.Attachments]

            print(f"\nProcessing: {subject}")

            folder_path = classify_email(
                subject=subject,
                body=body,
                sender_name=sender_name,
                sender_email=sender_email,
                to=to,
                cc=cc,
                attachments=attachments
            )
            print(f" -> Suggested folder: {folder_path}")


            # Normalize folder path: remove leading/trailing spaces, collapse multiple spaces, lowercase for check
            clean_path = folder_path.strip()
            # Remove multiple spaces and normalize slashes
            import re
            clean_path = re.sub(r"\\s+", " ", clean_path)
            # Only prepend if not already starting with 'postvak in/' (case-insensitive)
            if not clean_path.lower().startswith("postvak in/"):
                clean_path = f"Postvak In/{clean_path}"

            # Disallow weird folder names (e.g. 'Postvak In Opr', 'Postvak In ...', etc.)
            # Only allow folder paths that start with 'Postvak In/' and do not contain forbidden patterns
            forbidden_patterns = [r"postvak in[^/]*[^/]", r"\.\.\.", r"[^a-zA-Z0-9/ \-]", r"postvak in opr"]
            if not clean_path.lower().startswith("postvak in/") or any(re.search(p, clean_path.lower()) for p in forbidden_patterns):
                print(f"[ERROR] Suggested folder name '{folder_path}' is invalid. Using 'Postvak In/Unsorted' instead.")
                clean_path = "Postvak In/Unsorted"

            folder_path = clean_path

            # If CREATE_NEW_FOLDERS is False, always sort into Unsorted
            if not CREATE_NEW_FOLDERS:
                folder_path = f"{SOURCE_FOLDER_PATH}/Unsorted"
            target_folder = get_or_create_folder(source_folder, folder_path)
            # Check if target_folder is a Mail folder
            if hasattr(target_folder, 'DefaultItemType') and target_folder.DefaultItemType == 0:
                print(f"Moving mail to folder: {target_folder.Name}")
                mail.Move(target_folder)
                print("Moved successfully.")
                log_line = f"SORTED: '{subject}' -> '{folder_path}'"
                log_file.write(log_line + "\n")
            else:
                print(f"Target folder '{target_folder.Name}' is not a Mail folder. Cannot move mail.")
                log_line = f"SKIPPED: '{subject}' -> '{folder_path}' (Not a Mail folder)"
                log_file.write(log_line + "\n")
        except Exception as e:
            print(f"Error processing email: {e}")
            log_line = f"SKIPPED: '{getattr(mail, 'Subject', 'No Subject')}' -> ERROR: {e}"
            log_file.write(log_line + "\n")
        # Update progress bar only
        elapsed = time.time() - start_time
        remaining = (elapsed / idx) * (total_mails - idx) if idx > 0 else 0
        pbar.set_postfix({"Elapsed": f"{elapsed:.1f}s", "ETA": f"{remaining:.1f}s"})
        pbar.update(1)
    pbar.close()
    log_file.close()

if __name__ == "__main__":
    main()