import argparse
import os
import re
import sys
import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# ---------------------------
# Command-line arguments
# ---------------------------
parser = argparse.ArgumentParser(description="Book Metadata Scanner")

parser.add_argument(
    "--input", "-i",
    default="input.txt",
    help="Path to input TXT file (default: input.txt)"
)
parser.add_argument(
    "--output", "-o",
    default="books.xlsx",
    help="Path to output Excel file (default: books.xlsx)"
)
parser.add_argument(
    "--debug", "-d",
    default="debug_log.txt",
    help="Path to debug log file (default: debug_log.txt)"
)
parser.add_argument(
    "--logdir", "-l",
    default="logs_preview",
    help="Directory to store HTML/JSON previews (default: logs_preview)"
)

args = parser.parse_args()

# Default paths
INPUT_FILE = args.input
OUTPUT_FILE = args.output
DEBUG_FILE = args.debug
LOG_DIR = args.logdir
os.makedirs(LOG_DIR, exist_ok=True)

# Print defaults if no options provided
if len(sys.argv) == 1:
    print("No command-line options provided. Using defaults:")
    print(f"  Input file: {INPUT_FILE}")
    print(f"  Output file: {OUTPUT_FILE}")
    print(f"  Debug log: {DEBUG_FILE}")
    print(f"  Logs directory: {LOG_DIR}")
    print("\nStarting processing...\n")

# ---------------------------
# Read Google Books API Key
# ---------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
API_KEY_FILE = os.path.join(SCRIPT_DIR, "google_api.txt")
if os.path.exists(API_KEY_FILE):
    with open(API_KEY_FILE, "r", encoding="utf-8") as f:
        GOOGLE_BOOKS_API_KEY = f.read().strip()
else:
    print("Error: google_api.txt not found in script directory.")
    GOOGLE_BOOKS_API_KEY = None

# ---------------------------
# Helpers
# ---------------------------
def is_isbn(text):
    return re.fullmatch(r"\d{10}(\d{3})?", text.strip()) is not None

def clean_ocr_text(text):
    parts = text.split("-")
    title = parts[0].strip() if len(parts) > 0 else None
    author = parts[1].strip() if len(parts) > 1 else None
    return title, author

def save_preview(line_num, source, raw_content):
    path = os.path.join(LOG_DIR, f"log_{line_num}_{source}.txt")
    with open(path, "w", encoding="utf-8") as f:
        if isinstance(raw_content, (dict, list)):
            import json
            f.write(json.dumps(raw_content, indent=2, ensure_ascii=False))
        else:
            f.write(raw_content)

def log_debug(line_num, source, query, raw_content=None, error=None):
    with open(DEBUG_FILE, "a", encoding="utf-8") as f:
        f.write(f"\nLine {line_num} | Source: {source}\n")
        f.write(f"Query: {query}\n")
        if error:
            f.write(f"Error: {error}\n")
        if raw_content:
            preview_file = f"log_{line_num}_{source}.txt"
            f.write(f"Preview saved: {preview_file}\n")
    if raw_content:
        save_preview(line_num, source, raw_content)

# ---------------------------
# Fetchers
# ---------------------------
def fetch_by_google_books(title=None, author=None, line_num=0, original_text=None, isbn=None):
    if not GOOGLE_BOOKS_API_KEY:
        log_debug(line_num, "GoogleBooks", "No API key", error="Missing Google API key")
        return None
    try:
        if isbn:
            url = f"https://www.googleapis.com/books/v1/volumes?q=isbn:{isbn}&key={GOOGLE_BOOKS_API_KEY}"
        else:
            q = []
            if title:
                q.append(f"intitle:{title}")
            if author:
                q.append(f"inauthor:{author}")
            url = f"https://www.googleapis.com/books/v1/volumes?q={' '.join(q)}&key={GOOGLE_BOOKS_API_KEY}"

        r = requests.get(url, timeout=10)
        data = r.json()
        log_debug(line_num, "GoogleBooks", url, data)

        if "items" in data:
            info = data["items"][0]["volumeInfo"]
            return {
                "title": info.get("title"),
                "author": ", ".join(info.get("authors", [])) if info.get("authors") else None,
                "publisher": info.get("publisher"),
                "publishedDate": info.get("publishedDate"),
                "isbn": isbn,
                "language": info.get("language") or "Unknown",
                "source": "GoogleBooks"
            }
    except Exception as e:
        log_debug(line_num, "GoogleBooks", original_text, error=e)
    return None

def fetch_by_openlibrary(title=None, author=None, line_num=0, original_text=None, isbn=None):
    try:
        if isbn:
            url = f"https://openlibrary.org/isbn/{isbn}.json"
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                data = r.json()
                log_debug(line_num, "OpenLibraryISBN", url, data)
                return {
                    "title": data.get("title"),
                    "author": ", ".join([a["name"] for a in data.get("authors", [])]) if data.get("authors") else None,
                    "publisher": ", ".join(data.get("publishers", [])) if data.get("publishers") else None,
                    "publishedDate": data.get("publish_date"),
                    "isbn": isbn,
                    "language": ", ".join(data.get("languages", [])) if data.get("languages") else "Unknown",
                    "source": "OpenLibraryISBN"
                }
        else:
            q = []
            if title:
                q.append(f"title={title}")
            if author:
                q.append(f"author={author}")
            url = f"https://openlibrary.org/search.json?{'&'.join(q)}"
            r = requests.get(url, timeout=10)
            data = r.json()
            log_debug(line_num, "OpenLibrary", url, data)
            if "docs" in data and len(data["docs"]) > 0:
                doc = data["docs"][0]
                return {
                    "title": doc.get("title"),
                    "author": ", ".join(doc.get("author_name", [])) if doc.get("author_name") else None,
                    "publisher": ", ".join(doc.get("publisher", [])) if doc.get("publisher") else None,
                    "publishedDate": doc.get("first_publish_year"),
                    "isbn": doc.get("isbn", [None])[0],
                    "language": ", ".join(doc.get("language", [])) if doc.get("language") else "Unknown",
                    "source": "OpenLibrary"
                }
    except Exception as e:
        log_debug(line_num, "OpenLibrary", original_text, error=e)
    return None

def fetch_by_opac_sbn(title=None, author=None, line_num=0, original_text=None):
    try:
        url = f"https://opac.sbn.it/opacsbn/opaclib?db=solr_iccu&select_db=solr_iccu&searchForm=opac/iccu/free.jsp&resultForward=opac/iccu/full.jsp&do_cmd=search_show_cmd&format=xml&from=1&nentries=1&searchType=perfree&fname=none&value={title or ''}+{author or ''}"
        r = requests.get(url, timeout=10)
        log_debug(line_num, "OPAC", url, r.text)
        soup = BeautifulSoup(r.text, "html.parser")
        result = soup.find("title")
        if result:
            return {
                "title": result.text.strip(),
                "author": author,
                "publisher": "Unknown",
                "publishedDate": "Unknown",
                "isbn": None,
                "language": "Unknown",
                "source": "OPAC SBN"
            }
    except Exception as e:
        log_debug(line_num, "OPAC", original_text, error=e)
    return None

def fetch_from_amazon(query, line_num=0):
    try:
        url = f"https://www.amazon.it/s?k={query}"
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(url, headers=headers, timeout=10)
        log_debug(line_num, "Amazon", url, r.text)
        soup = BeautifulSoup(r.text, "html.parser")

        title = None
        author = None

        title_tag = soup.select_one("h2 a span")
        if title_tag:
            title = title_tag.text.strip()

        author_tag = soup.select_one(".a-color-secondary .a-size-base")
        if author_tag:
            author = author_tag.text.strip()

        return {
            "title": title,
            "author": author,
            "publisher": None,
            "publishedDate": None,
            "isbn": None,
            "language": "Unknown",
            "source": "Amazon"
        }
    except Exception as e:
        log_debug(line_num, "Amazon", query, error=e)
    return None

# ---------------------------
# Main
# ---------------------------
def main():
    wb = Workbook()
    ws = wb.active
    ws.append(["Title", "Author", "Publisher", "Publication Date", "ISBN", "Language", "Source"])

    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip()]

    total_lines = len(lines)
    start_time = time.time()

    for idx, line in enumerate(lines, 1):
        try:
            timestamp, data = line.split("\t", 1)
        except ValueError:
            print(f"[{idx}/{total_lines}] Skipping invalid line: {line}")
            continue

        is_isbn_flag = is_isbn(data)
        data_type = "ISBN" if is_isbn_flag else "OCR Text"

        elapsed = time.time() - start_time
        avg_time = elapsed / idx
        remaining_time = avg_time * (total_lines - idx)
        eta = time.strftime("%H:%M:%S", time.gmtime(remaining_time))

        print(f"[{idx}/{total_lines}] Processing: '{data}' ({data_type}) | ETA: {eta}")
        sys.stdout.flush()

        title_guess, author_guess = clean_ocr_text(data)

        book_info = None
        if is_isbn_flag:
            book_info = fetch_by_google_books
