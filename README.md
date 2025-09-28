
# Book Metadata Scanner

A Python script to scan a list of books (ISBN or OCR text from the front page) and retrieve detailed metadata including **Title, Author, Publisher, Publication Date, ISBN, Language, and Source**.  

The script searches multiple sources:

- **Google Books API**  
- **OpenLibrary API**  
- **OPAC SBN** (Italian public library catalog)  
- **Amazon** (as a fallback, web scraping)

It logs detailed debug information and saves HTML/JSON previews for every query.

---

## Features

- Supports **Italian, UK, and US books**.  
- Detects **ISBN** automatically; otherwise, uses OCR text.  
- **Single-threaded** for reliability and simplicity.  
- Creates an **Excel file** with all metadata.  
- Maintains a **debug log** and HTML/JSON previews for each query.  
- Adds **book language** in the Excel output.  
- Allows **custom input/output paths** via command line arguments.  
- Reads **Google Books API key** from an external `google_api.txt` file.  

---

## Installation

1. Clone this repository:

```bash
git clone https://github.com/yourusername/book-metadata-scanner.git
cd book-metadata-scanner
```

2. Create a Python virtual environment (optional but recommended):

```bash
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows
```

3. Install required packages:

```bash
pip install requests beautifulsoup4 openpyxl
```

4. Create a file named `google_api.txt` in the same directory as the script and paste your **Google Books API key** inside:

```
YOUR_GOOGLE_BOOKS_API_KEY
```

---

## Usage

### Input file format

The input TXT file should contain one book per line, in the format:

```
timestamp<TAB>ISBN
timestamp<TAB>OCR text from book front page
```

Example:

```
2025-09-27 10:00:00    9788817163835
2025-09-27 10:05:00    Il Nome della Rosa - Umberto Eco
```

---

### Running the script

With defaults (input: `input.txt`, output: `books.xlsx`, debug log: `debug_log.txt`, previews: `logs_preview/`):

```bash
python scan_books.py
```

With custom paths:

```bash
python scan_books.py --input mybooks.txt --output results.xlsx --debug mylog.txt --logdir previews
```

---

## Output

After completion you will get:

- **Excel output** → `books.xlsx` (or custom path)  
- **Debug log** → `debug_log.txt` (or custom path)  
- **HTML/JSON previews** → `logs_preview/` (or custom directory)  

### Excel columns:

| Title | Author | Publisher | Publication Date | ISBN | Language | Source |
|-------|--------|-----------|-----------------|------|----------|--------|

- **Source** indicates which service provided the metadata.  
- **Language** is taken from APIs (if available).  

---

## Notes

- Amazon scraping is **used only if other sources fail**.  
- The script is **single-threaded** to reduce errors and avoid being blocked by Amazon.  
- OCR text parsing is basic; results improve if the input text is clean.  
- The **Google Books API key** must be placed in `google_api.txt` in the same directory as the script.  

---

## License

This project is licensed under the MIT License.  

---

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.
