import requests
import openpyxl
import re
import time
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
import sys
import json
import hashlib
import os

# ----------------------------
# CONFIGURATION
# ----------------------------
INPUT_FILE = "scan_results.txt"
OUTPUT_FILE = "books.xlsx"

GOOGLE_BOOKS_API_KEY = ""

DEBUG_FILE = "debug_log.txt"
LOG_DIR = "logs_preview"  # directory to save html/json previews

AMAZON_SITES = [
    "https://www.amazon.it/s?k=",
    "https://www.amazon.co.uk/s?k=",
    "https://www.amazon.com/s?k="
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"
}

os.makedirs(LOG_DIR, exist_ok=True)

# ----------------------------
# HELPER FUNCTIONS
# ----------------------------
def is_isbn(text):
    clean_text = text.replace("-", "").replace(" ", "")
    return clean_text.isdigit() and len(clean_text) in (10, 13)

def clean_ocr_text(text):
    text = re.sub(r'\s+', ' ', text).strip()
    if " - " in text:
        title, author = text.split(" - ", 1)
    elif "\n" in text:
        parts = text.split("\n")
        title = parts[0]
        author = parts[1] if len(parts) > 1 else ""
    else:
        words = text.split(" ")
        mid = len(words) // 2
        title = " ".join(words[:mid])
        author = " ".join(words[mid:])
    return title.strip(), author.strip()

def safe_filename(text):
    if is_isbn(text):
        return text
    else:
        return hashlib.md5(text.encode("utf-8")).hexdigest()

def save_preview(text, content):
    filename = os.path.join(LOG_DIR, f"log_{safe_filename(text)}.txt")
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content[:1000])

def log_debug(line_num, source, query, result, error=None):
    entry = {
        "line": line_num,
        "source": source,
        "query": query,
        "result": result,
        "error": str(error) if error else None
    }
    with open(DEBUG_FILE, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")

# ----------------------------
# METADATA SOURCES
# ----------------------------
def fetch_by_google_books(title=None, author=None, line_num=None, original_text=None, isbn=None):
    query = f"isbn:{isbn}" if isbn else ""
    if not isbn:
        if title:
            query += f"intitle:{title}"
        if author:
            query += f"+inauthor:{author}"
    url = f"https://www.googleapis.com/books/v1/volumes?q={query}&key={GOOGLE_BOOKS_API_KEY}"
    try:
        response = requests.get(url, timeout=10)
        save_preview(original_text or query, response.text)
        items = response.json().get("items", [])
        result = None
        if items:
            info = items[0]["volumeInfo"]
            result = {
                "title": info.get("title"),
                "author": ", ".join(info.get("authors", [])) if info.get("authors") else None,
                "publisher": info.get("publisher"),
                "publishedDate": info.get("publishedDate"),
                "isbn": isbn,
                "language": info.get("language") or "Unknown",
                "source": "GoogleBooks"
            }
        log_debug(line_num, "GoogleBooks", query, result)
        return result
    except Exception as e:
        log_debug(line_num, "GoogleBooks", query, None, error=e)
        return None

def fetch_by_openlibrary(title=None, author=None, line_num=None, original_text=None, isbn=None):
    if isbn:
        url = f"https://openlibrary.org/api/books?bibkeys=ISBN:{isbn}&format=json&jscmd=data"
        try:
            response = requests.get(url, timeout=10)
            save_preview(original_text or isbn, response.text)
            data = response.json()
            if data:
                info = next(iter(data.values()))
                lang_list = [l['key'].split('/')[-1] for l in info.get("languages", [])] if info.get("languages") else []
                result = {
                    "title": info.get("title"),
                    "author": ", ".join([a['name'] for a in info.get("authors", [])]) if info.get("authors") else None,
                    "publisher": ", ".join(info.get("publishers", [])) if info.get("publishers") else None,
                    "publishedDate": info.get("publish_date"),
                    "isbn": isbn,
                    "language": ", ".join(lang_list) if lang_list else "Unknown",
                    "source": "OpenLibraryISBN"
                }
                log_debug(line_num, "OpenLibrary_ISBN", isbn, result)
                return result
            return None
        except Exception as e:
            log_debug(line_num, "OpenLibrary_ISBN", isbn, None, error=e)
            return None
    else:
        query = ""
        if title:
            query += f"title={quote_plus(title)}"
        if author:
            query += f"&author={quote_plus(author)}"
        url = f"https://openlibrary.org/search.json?{query}"
        try:
            response = requests.get(url, timeout=10)
            save_preview(original_text or query, response.text)
            docs = response.json().get("docs", [])
            result = None
            if docs:
                doc = docs[0]
                result = {
                    "title": doc.get("title"),
                    "author": ", ".join(doc.get("author_name", [])) if doc.get("author_name") else None,
                    "publisher": ", ".join(doc.get("publisher", [])) if doc.get("publisher") else None,
                    "publishedDate": doc.get("first_publish_year"),
                    "isbn": doc.get("isbn", [None])[0],
                    "language": ", ".join(doc.get("language", [])) if doc.get("language") else "Unknown",
                    "source": "OpenLibrary"
                }
            log_debug(line_num, "OpenLibrary", query, result)
            return result
        except Exception as e:
            log_debug(line_num, "OpenLibrary", query, None, error=e)
            return None

def fetch_by_opac_sbn(title=None, author=None, line_num=None, original_text=None):
    query = f"{title} {author}" if title and author else title or author
    url = f"https://opac.sbn.it/risultati-ricerca?f[title]={quote_plus(query)}"
    try:
        response = requests.get(url, timeout=10)
        save_preview(original_text or query, response.text)
        soup = BeautifulSoup(response.text, "html.parser")
        first_result = soup.select_one(".title a")
        result = None
        if first_result:
            result = {
                "title": first_result.text.strip(),
                "author": "Unknown",
                "publisher": "Unknown",
                "publishedDate": "Unknown",
                "isbn": None,
                "language": "Unknown",
                "source": "OPAC SBN"
            }
        log_debug(line_num, "OPAC SBN", query, result)
        return result
    except Exception as e:
        log_debug(line_num, "OPAC SBN", query, None, error=e)
        return None

def fetch_from_amazon(text, line_num=None):
    result = None
    for site in AMAZON_SITES:
        search_url = f"{site}{quote_plus(text)}"
        try:
            response = requests.get(search_url, headers=HEADERS, timeout=10)
            save_preview(text, response.text)
            soup = BeautifulSoup(response.text, "html.parser")
            first_result = soup.select_one("h2 a.a-link-normal")
            if not first_result:
                log_debug(line_num, "Amazon_SearchHTML", search_url, {"error": "No first_result found"})
                continue

            title = first_result.text.strip()
            book_url = "https://www.amazon.it" + first_result['href']
            book_resp = requests.get(book_url, headers=HEADERS, timeout=10)
            save_preview(text, book_resp.text)
            book_soup = BeautifulSoup(book_resp.text, "html.parser")

            author_tag = book_soup.select_one(".author a")
            author = author_tag.text.strip() if author_tag else "Unknown"

            publisher = "Unknown"
            published_date = "Unknown"
            isbn = None
            detail_items = book_soup.select("#detailBullets_feature_div li")
            for li in detail_items:
                key = li.select_one("span.a-text-bold")
                value = li.select_one("span.a-size-base")
                if key and value:
                    key_text = key.text.strip()
                    value_text = value.text.strip()
                    if "Publisher" in key_text:
                        publisher = value_text
                        year_match = re.search(r"\d{4}", publisher)
                        if year_match:
                            published_date = year_match.group(0)
                    if "ISBN-13" in key_text:
                        isbn_match = re.search(r"\d{10,13}", value_text)
                        if isbn_match:
                            isbn = isbn_match.group(0)

            result = {
                "title": title,
                "author": author,
                "publisher": publisher,
                "publishedDate": published_date,
                "isbn": isbn,
                "language": "Unknown",
                "source": "Amazon"
            }
            log_debug(line_num, "Amazon_Result", book_url, result)
            return result
        except Exception as e:
            log_debug(line_num, "Amazon_Exception", search_url, None, error=e)
    return None

# ----------------------------
# MAIN SCRIPT
# ----------------------------
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Books"
ws.append(["Title", "Author", "Publisher", "Publication Date", "ISBN", "Language", "Source"])

open(DEBUG_FILE, "w").close()  # clear debug file

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

    # --------------------------
    # Optimized search sequence
    # --------------------------
    book_info = None
    if is_isbn_flag:
        book_info = fetch_by_google_books(isbn=data, line_num=idx, original_text=data)
        if not book_info:
            book_info = fetch_by_openlibrary(isbn=data, line_num=idx, original_text=data)
    else:
        book_info = fetch_by_google_books(title=title_guess, author=author_guess, line_num=idx, original_text=data)
        if not book_info:
            book_info = fetch_by_openlibrary(title=title_guess, author=author_guess, line_num=idx, original_text=data)
        if not book_info:
            book_info = fetch_by_opac_sbn(title=title_guess, author=author_guess, line_num=idx, original_text=data)

    # Only query Amazon sequentially if all above fail
    if not book_info:
        book_info = fetch_from_amazon(data, line_num=idx)

    # Append to Excel
    if book_info:
        ws.append([
            book_info.get("title") or "Unknown",
            book_info.get("author") or "Unknown",
            book_info.get("publisher") or "Unknown",
            book_info.get("publishedDate") or "Unknown",
            book_info.get("isbn") or "",
            book_info.get("language") or "Unknown",
            book_info.get("source") or ""
        ])
    else:
        ws.append(["Unknown", "Unknown", "Unknown", "Unknown", "", "Unknown", "Not Found"])

    time.sleep(1)

wb.save(OUTPUT_FILE)
print(f"\nExcel file saved as {OUTPUT_FILE}")
print(f"Debug log saved as {DEBUG_FILE}")
print(f"HTML/JSON previews saved in folder: {LOG_DIR}")
