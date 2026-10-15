# Resume Processing Automation Tool

A Python-based automation script designed to extract key contact and professional information from PDF and Word (`.docx`, `.doc`) resume files, parse section headers intelligently, and compile the structured data into a cleanly formatted Excel spreadsheet with tracking logs.

---

## 🚀 Features

* **Multi-Format Support**: Automatically detects and extracts text from both PDF documents (via `pdfplumber`) and Word documents (via `python-docx`).
* **Intelligent Information Extraction**: 
  * Uses regular expressions to capture phone numbers and email addresses.
  * Employs heuristic search logic to identify candidate names from the top lines of text.
  * Parses distinct resume sections—such as Education, Technical Skills, Experience, and Projects—even when documents use varied header naming conventions.
* **Deduplication & Tracking**: Maintains a `processed_log.txt` file to ensure resumes are only parsed once, preventing duplicate entries in your output sheet.
* **Excel Formatting & Auto-Styling**: Compiles data into an Excel workbook using `pandas` and `openpyxl`, automatically applying grid filters, fixed column widths, generous row heights, and text wrapping for easy readability.

---

## 🛠️ Prerequisites & Dependencies

The script relies on several external Python libraries. You can install them using `pip`:

```bash
pip install pdfplumber python-docx pandas openpyxl
```

* **Python Libraries Used**:
  * `os`, `re`, `pathlib` (Standard Library)
  * `pdfplumber`
  * `docx` (`python-docx`)
  * `pandas`
  * `openpyxl`

---

## 📂 Project Structure

```text
├── automation_code.py      # Main automation script
├── Resumes/                # Default directory containing source resume files (PDF/DOCX)
├── output.xlsx             # Compiled and styled Excel output sheet
└── processed_log.txt       # Tracking log of already processed files
```

---

## ⚙️ How It Works

1. **Log Loading**: Reads `processed_log.txt` to compile a set of previously handled filenames.
2. **Directory Scanning**: Iterates through all files inside the `Resumes/` folder.
3. **Text Extraction & Cleaning**: Extracts raw strings from supported file types and cleans up carriage returns, tabs, excessive whitespace, and decorative characters.
4. **Parsing**: Runs regex patterns and keyword matchers to extract:
   * Candidate Name, Email, and Phone
   * Education, Skills, Experience, and Projects sections
5. **Excel Export & Formatting**: Appends new rows to `output.xlsx` (creating it if it doesn't exist) and formats cell alignments, dimensions, and filters.
6. **Log Update**: Marks successfully processed files in `processed_log.txt`.

---

## 🏃‍♂️ Usage

1. Place your target resume files (`.pdf`, `.docx`) into a folder named `Resumes` in the same directory as the script (or modify the `resume_folder` parameter in `process_resumes()`).
2. Run the script from your terminal:

```bash
python automation_code.py
```

3. View the generated results in `output.xlsx`.
