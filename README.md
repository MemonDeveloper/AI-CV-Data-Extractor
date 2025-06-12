# AI CV Data Extractor and Renamer 📄✨

This Python tool automates the process of extracting key information from CVs (PDF, DOC, and DOCX formats) using the Gemini AI model and OCR capabilities. It then organizes this data into an Excel spreadsheet and can even rename PDF CVs based on the extracted name.

## Features 🚀

  * **Multi-format Support**: Processes PDF, DOC, and DOCX files.
  * **Intelligent Data Extraction**: Leverages the Gemini 2.0 Flash model to accurately pull out details like name, contact information, education, and work history.
  * **OCR Fallback**: Automatically uses OCR (Tesseract) for scanned PDFs or images within PDFs when direct text extraction isn't possible.
  * **PDF Renaming**: Renames PDF files to the candidate's name (e.g., "John Doe.pdf") for easier organization.
  * **Excel Output**: Compiles all extracted data into a clean, well-formatted Excel file (`.xlsx`).
  * **Rate Limit Handling**: Includes a built-in delay to respect API rate limits.
  * **Error Handling**: Gracefully handles errors during file processing and API calls.

-----

## Prerequisites 🛠️

Before running this tool, ensure you have the following installed and configured:

1.  **Python 3.x**: Download from [python.org](https://www.python.org/downloads/).
2.  **Tesseract OCR**:
      * **Windows**: Download the installer from [UB Mannheim](https://www.google.com/search?q=https://github.com/UB-Mannheim/tesseract/wiki).
      * **macOS**: Install via Homebrew: `brew install tesseract`
      * **Linux**: Install via your package manager (e.g., `sudo apt install tesseract-ocr` for Debian/Ubuntu).
      * **Important**: Note the installation path for Tesseract. You'll need to update `pytesseract.pytesseract.tesseract_cmd` in the script.
3.  **Microsoft Word (for `.doc` files)**: The script uses the COM interface for `.doc` files, which requires a local installation of Microsoft Word.
4.  **Google Gemini API Key**:
      * Obtain one from [Google AI Studio](https://aistudio.google.com/app/apikey).
      * Replace `"AIzaS"` with your actual API key in the `genai.configure(api_key="YOUR_API_KEY")` line.

-----

## Installation 💻

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/MemonDeveloper/AI-CV-Data-Extractor.git
    cd AI-CV-Data-Extractor
    ```
2.  **Install Python dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
    (You'll need to create a `requirements.txt` file, see the "Creating `requirements.txt`" section below.)

-----

## Creating `requirements.txt` 📝

Create a file named `requirements.txt` in the root directory of your project and paste the following dependencies:

```
google-generativeai
pandas
PyMuPDF
pytesseract
python-docx
pdf2image
demjson3
pypiwin32  # For Windows users, for the win32com.client part
```

-----

## Configuration ⚙️

Open the `main.py` file and modify the following lines:

  * **Gemini API Key**:
    ```python
    genai.configure(api_key="YOUR_API_KEY") # 🔁 Replace this with your real API key
    ```
  * **Tesseract OCR Path**:
    ```python
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' # 🔁 Modify if needed
    ```
    (Adjust this path to where Tesseract is installed on your system. For macOS/Linux, it might just be `'tesseract'` if it's in your PATH, or `/usr/local/bin/tesseract`, etc.)

-----

## Usage ▶️

1.  **Place your CVs**: Put all the CV files (PDF, DOC, DOCX) you want to process into a single folder.
2.  **Run the script**:
    ```bash
    python main.py
    ```
3.  **Enter folder path**: The script will prompt you to enter the full path to the folder containing your CVs. For example:
    ```
    Enter the full folder path containing CV files: C:\Users\YourUser\Documents\CVs
    ```
    Or on Linux/macOS:
    ```
    Enter the full folder path containing CV files: /home/youruser/Documents/CVs
    ```

The script will then:

  * Iterate through each supported file in the specified folder.
  * Extract text from the CV.
  * Send the text to the Gemini API for structured data extraction.
  * Rename PDF files based on the extracted "Name" field (if available and a PDF).
  * Compile all extracted data into an Excel file named `output_YYYY-MM-DD_HH-MM.xlsx` in the same directory where you run the script.

### Output CSV/Excel Format 📊

The generated Excel file will have the following columns in this exact order:

`InsertDate`, `S.no`, `Name`, `City`, `DOB`, `Age`, `Gender`, `Last/currentQualification`, `Last/currentInstitute`, `Last/currentDegreeYear`, `Mobile No.`, `Email`, `Last/current Company`, `Last/currentCompanyPosition`, `Last/currentStatus`

-----

## Important Notes ⚠️

  * **API Key Security**: **Never commit your actual API key directly to a public GitHub repository.** For production environments, consider using environment variables to store your API key.
  * **Microsoft Word Requirement**: The `.doc` file processing relies on `win32com.client`, which is a Windows-specific library that interacts with Microsoft Word. This functionality will not work on macOS or Linux without a Windows environment (e.g., Wine or a VM).
  * **Rate Limits**: The Gemini API has rate limits. The script includes a `time.sleep()` to pause between requests to prevent hitting these limits. If you have a large number of CVs, the process may take some time.
  * **OCR Accuracy**: The accuracy of OCR can vary depending on the quality of the scanned documents.
  * **Gemini Model Performance**: The quality of the extracted data depends heavily on the Gemini model's ability to interpret the CV text. While generally very good, complex or unconventional CV formats might yield less accurate results.
  * **Filename Sanitization**: The script sanitizes filenames to remove invalid characters, ensuring compatibility across different operating systems.

-----

## Contributing 🤝

Feel free to fork this repository, open issues, and submit pull requests if you have suggestions for improvements or bug fixes.

-----
