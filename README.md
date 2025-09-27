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

---

## Installation

1. Clone this repository:

```bash
git clone https://github.com/yourusername/book-metadata-scanner.git
cd book-metadata-scanner
```

2. Add your own Google API


