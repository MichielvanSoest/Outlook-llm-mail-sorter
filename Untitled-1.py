import requests
import pdfplumber

def ask_pdf(question: str, pdf_path: str, model_id: str = "google/gemma-3-12b", max_tokens: int = 1000):
    """
    Extracts text from a PDF preserving layout, and sends a prompt to LM Studio API.
    
    Args:
        question (str): The question you want to ask about the PDF.
        pdf_path (str): Path to the PDF file.
        model_id (str): LM Studio model ID to use.
        max_tokens (int): Maximum number of tokens in the response.
    
    Returns:
        str: The model's answer.
    """
    
    # 1. Extract PDF text preserving layout
    pdf_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text(x_tolerance=2)  # smaller tolerance preserves layout
            if page_text:
                pdf_text += f"--- Page {page_number} ---\n{page_text}\n"
    
    # Optional: truncate if PDF is very large
    if len(pdf_text) > 15000:
        pdf_text = pdf_text[:15000] + "\n...[truncated]...\n"
    
    # 2. Create prompt combining PDF text and question
    prompt = f"Document content:\n{pdf_text}\n\nQuestion: {question}\nAnswer:"
    
    # 3. Send to LM Studio API
    url = "http://localhost:1234/v1/completions"
    data = {
        "model": model_id,
        "prompt": prompt,
        "max_tokens": max_tokens
    }
    
    # Do the actual api call using above url and data (json format)
    response = requests.post(url, json=data)
    
    if response.status_code == 200:
        result = response.json()
        return result["choices"][0]["text"]
    else:
        raise RuntimeError(f"Error {response.status_code}: {response.text}")


# =======================
# Example usage
# =======================

question = r"Lees deze factuur en haal alle relevante gegevens eruit die nodig zijn om deze in een CRM-systeem te registreren. Geef uitsluitend een JSON-object als output, zonder extra uitleg of tekst. Neem het volgende mee in de JSON: • factuurnummer • factuurdatum • vervaldatum • leverancier: naam, adres, contactgegevens • klant: naam, adres, contactgegevens (indien aanwezig) • artikelen of diensten: omschrijving, aantal, prijs per eenheid, totaal per regel • subtotaal • btw: tarief per regel, totaal btw-bedrag • totaalbedrag • betaalwijze • referentienummers of opmerkingen Als er tabellen aanwezig zijn, neem alle rijen mee in de JSON."
pdf_path = r"D:\Downloads\factuur-voorbeeld-zzp.pdf"

answer = ask_pdf(question, pdf_path)
print("Answer from LM Studio:")
print(answer)
