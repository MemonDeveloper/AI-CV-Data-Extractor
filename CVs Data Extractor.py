import os
import fitz  # PyMuPDF
import pytesseract
import datetime
import google.generativeai as genai
import pandas as pd
import demjson3
import re
import time
import docx
import pythoncom
import win32com.client
import string
from pdf2image import convert_from_path

# Gemini API Configuration
genai.configure(api_key="ddddd")  # 🔁 Replace this with your real API key
model = genai.GenerativeModel("models/gemini-2.0-flash")

# Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # 🔁 Modify if needed

# DOC file reader (using MS Word COM interface)
def doc_to_text(doc_path):
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    text = ""
    try:
        doc = word.Documents.Open(doc_path)
        text = doc.Content.Text
        doc.Close()
    except Exception as e:
        text = f"Error reading {doc_path}: {e}"
    finally:
        word.Quit()
    return text

# Extract text from PDF/DOC/DOCX
def extract_text_from_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    text = ""

    if ext == ".pdf":
        try:
            doc = fitz.open(filepath)
            for page in doc:
                page_text = page.get_text().strip()
                if page_text:
                    text += page_text + "\n"
            doc.close()

            if not text.strip():
                print("🔍 No direct text found, using OCR fallback...")
                images = convert_from_path(filepath)
                for img in images:
                    text += pytesseract.image_to_string(img, lang='eng') + "\n"
        except Exception as e:
            text = f"Error processing {filepath}: {e}"

    elif ext == ".docx":
        try:
            doc = docx.Document(filepath)
            for para in doc.paragraphs:
                if para.text.strip():
                    text += para.text.strip() + "\n"
        except Exception as e:
            text = f"Error processing {filepath}: {e}"

    elif ext == ".doc":
        text = doc_to_text(filepath)

    else:
        text = f"Unsupported file format: {ext}"

    return text

# Clean code block formatting if needed
def clean_code_block(text):
    text = text.strip()
    if text.startswith("```") and text.endswith("```"):
        text = "\n".join(text.splitlines()[1:-1])
    return text

# Convert response text to dictionary
def extract_dict_from_text(text):
    try:
        match = re.search(r"\{[\s\S]*\}", text)
        if not match:
            raise ValueError("No dictionary-like object found in response.")
        dict_text = clean_code_block(match.group(0))
        return demjson3.decode(dict_text)
    except Exception as e:
        raise ValueError(f"Failed to parse dictionary: {e}")

# Build prompt and call Gemini API
def extract_details_with_gemini(text, serial_no):
    prompt = f"""
You are an expert in reading CVs. Extract the following fields from the CV text below:

InsertDate, S.no, Name, City, DOB, Age, Gender, Last/currentQualification,
Last/currentInstitute, Last/currentDegreeYear, Mobile No., Email,
Last/current Company, Last/currentCompanyPosition, Last/currentStatus

Note:
- For DOB, Date Format ([$-en-US]d-mmm-yy;@)
- For Last/currentCompanyPosition, return the position or job title held in the last/current company.
- For Last/currentDegreeYear, return the year the degree was completed (e.g., "2019") or "Continue" if the degree is ongoing.
- For Last/currentStatus, return "On Job" if the person is currently employed, otherwise "Jobless".
- Mobile No. and Email can be multiple. Return them as lists of strings, e.g.:
  Mobile No.: +92 300-2485177, +92 310-6510318
  Email: email1@example.com, email2@example.com

InsertDate = {datetime.date.today().isoformat()}
S.no = {serial_no}

CV Text:
{text}

Respond only with a valid JSON object or Python dictionary. Do not add any explanation or markdown. Strictly return the dictionary or JSON only.
"""
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Error from Gemini: {e}"

# Gemini Flash: Max 15 requests per minute
REQUEST_LIMIT = 15
TIME_WINDOW = 60
MIN_DELAY = TIME_WINDOW / REQUEST_LIMIT

# Filename sanitizer
def sanitize_filename(name):
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    sanitized = ''.join(c for c in name if c in valid_chars)
    return sanitized.replace(' ', ' ')  # spaces replaced with underscores

# Generate unique filename if file exists (adds (1), (2), etc.)
def get_unique_filename(folder, base_name, ext=".pdf"):
    sanitized_base = sanitize_filename(base_name)
    candidate = f"{sanitized_base}{ext}"
    counter = 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{sanitized_base} ({counter}){ext}"
        counter += 1
    return candidate

# Main execution function
def main():
    folder_path = input("Enter the full folder path containing CV files: ").strip()
    if not os.path.isdir(folder_path):
        print(f"❌ Error: Directory '{folder_path}' does not exist.")
        return

    all_data = []
    serial = 1

    for filename in os.listdir(folder_path):
        if filename.lower().endswith((".pdf", ".doc", ".docx")):
            full_path = os.path.join(folder_path, filename)
            print(f"\n📄 Processing: {filename}")

            text = extract_text_from_file(full_path)
            result = extract_details_with_gemini(text, serial)

            try:
                data_dict = extract_dict_from_text(result)
                all_data.append(data_dict)

                # Rename PDF if extension is .pdf and Name is present
                if filename.lower().endswith(".pdf") and 'Name' in data_dict:
                    extracted_name = data_dict['Name']
                    if extracted_name and isinstance(extracted_name, str) and extracted_name.strip():
                        # Proper case conversion (title case)
                        proper_name = extracted_name.title()
                        new_filename = get_unique_filename(folder_path, proper_name, ext=".pdf")
                        new_full_path = os.path.join(folder_path, new_filename)
                        os.rename(full_path, new_full_path)
                        print(f"🔄 Renamed '{filename}' to '{new_filename}'")

            except Exception as e:
                print(f"⚠️ Error parsing response for {filename}: {e}")
                all_data.append({
                    "InsertDate": str(datetime.date.today()),
                    "S.no": serial,
                    "Error": f"Parse failed: {e}"
                })

            serial += 1
            print(f"⏳ Sleeping for {MIN_DELAY:.2f} seconds to respect rate limit...")
            time.sleep(MIN_DELAY)

    if all_data:
        df = pd.DataFrame(all_data)
        df = df.fillna("-")
        df = df.replace(r"^\s*$", "-", regex=True)

        # Normalize string fields (convert to title case)
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].apply(lambda x: x.title() if isinstance(x, str) else x)

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
        output_filename = f"output_{timestamp}.xlsx"
        df.to_excel(output_filename, index=False)
        print(f"\n✅ All data saved to {output_filename}")
    else:
        print("⚠️ No data processed.")

if __name__ == "__main__":
    main()
