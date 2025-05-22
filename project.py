#!/usr/bin/env python3
"""
A utility script to ingest text from various document formats, extract people’s names and interests
using the Deepseek API, and maintain a recency-ordered CSV of results.

Steps:
 1. Load environment variables (including DEEPSEEK_API_KEY).
 2. Load existing records from the CSV into an OrderedDict to preserve order.
 3. For each input file:
    - Validate extension (.docx, .pdf, .pages, .txt).
    - Extract plain text from the document.
    - Call Deepseek to extract all names and interests as a JSON array.
    - Upsert each (filename, name) entry, merging interests and moving to the end to record recency.
 4. Rewrite the CSV in the updated recency order.

Usage:
  python project.py file1.docx file2.pdf notes.txt --out results.csv

  Description:
    - Provide one or more input files (.docx, .pdf, .pages, .txt).
    - Use `-o` or `--out` to set the CSV output path (default: `results.csv`).
    - Ensure the `DEEPSEEK_API_KEY` environment variable is set or loaded via a `.env` file (using python-dotenv).

  Examples:
    # Basic usage with default output
    python project.py samples.docx

    # Using a custom output file
    python project.py report.txt summary.pages -o people.csv

"""

import os
import zipfile
import tempfile
import json
import csv
import argparse
import requests

from collections import OrderedDict
from docx import Document
import PyPDF2
from dotenv import load_dotenv

# ——— Configuration ———
load_dotenv()
DEEPSEEK_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"
EXTS = ("docx", "pdf", "pages", "txt")
CSV_FIELDS = ["filename", "name", "interests"]


# ——— Text extraction functions ———
def extract_text_from_docx(path):
    """
    Extracts and concatenates all paragraphs from a .docx file.
    """
    doc = Document(path)
    return "\n\n".join(p.text for p in doc.paragraphs if p.text)


def extract_text_from_pdf(path):
    """
    Extracts and concatenates text from all pages of a PDF.
    """
    reader = PyPDF2.PdfReader(path)
    return "\n\n".join(page.extract_text() or "" for page in reader.pages)


def extract_text_from_pages(path):
    """
    Extracts text from a .pages file by locating and reading its embedded PDF preview.
    """
    with zipfile.ZipFile(path, "r") as z:
        cands = [
            f for f in z.namelist() if f.startswith("QuickLook/") and f.endswith(".pdf")
        ]
        if not cands:
            raise FileNotFoundError("No QuickLook/Preview.pdf in .pages")
        with z.open(cands[0]) as embedded, tempfile.NamedTemporaryFile(
            delete=False, suffix=".pdf"
        ) as tf:
            tf.write(embedded.read())
            return extract_text_from_pdf(tf.name)


def extract_text_from_txt(path):
    """
    Reads and returns the entire contents of a plain text (.txt) file.
    """
    with open(path, encoding="utf-8") as f:
        return f.read()


# ——— Deepseek analysis function ———
def analyze_text_with_deepseek(text):
    """
    Sends the given text to the Deepseek API and returns a list of
    {"name": ..., "interests": [...]} objects parsed from the response JSON.
    """

    prompt = (
        "Extract *all* people’s names and interests from the text below.\n\n"
        "Respond ONLY with a JSON array of objects, each having:\n"
        '  • "name": string\n'
        '  • "interests": array of strings\n\n'
        "Example:\n"
        '[ {"name":"Alice","interests":["x","y"]}, … ]\n\n'
        f"Text:\n```{text}```"
    )
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0,
    }
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_KEY}",
        "Content-Type": "application/json",
    }
    r = requests.post(DEEPSEEK_URL, json=payload, headers=headers)
    r.raise_for_status()
    raw = r.json()["choices"][0]["message"]["content"].strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        import re

        m = re.search(r"(\[.*\])", raw, re.DOTALL)
        if m:
            return json.loads(m.group(1))
        raise ValueError(f"Could not parse JSON:\n{raw}")


# ——— CSV load and write utilities ———
def load_csv_into_ordereddict(path):
    """
    Loads existing CSV data into an OrderedDict mapping (filename, name) -> interests list.
    Preserves insertion order for recency tracking.
    """
    od = OrderedDict()
    if not os.path.isfile(path):
        return od
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            key = (row["filename"], row["name"])
            try:
                interests = json.loads(row["interests"])
            except:
                interests = []
            od[key] = interests
    return od


def write_ordereddict_to_csv(path, od):
    """
    Writes the OrderedDict contents back to a CSV file, preserving the current order.
    """
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        w.writeheader()
        for (fn, name), interests in od.items():
            w.writerow(
                {
                    "filename": fn,
                    "name": name,
                    "interests": json.dumps(interests, ensure_ascii=False),
                }
            )


# ——— Main orchestration ———
def main():
    print(
        "Starting processing of input files. For large documents, this may take a while…"
    )
    p = argparse.ArgumentParser()
    p.add_argument("files", nargs="+", help="(.docx,.pdf,.pages,.txt)")
    p.add_argument("-o", "--out", default="results.csv")
    args = p.parse_args()

    # 1) Load existing into OrderedDict (preserves original order)
    od = load_csv_into_ordereddict(args.out)

    # 2) Process each input file
    for path in args.files:
        if not os.path.isfile(path):
            print(f"⚠️ Skipping missing: {path}")
            continue
        ext = path.lower().rsplit(".", 1)[-1]
        if ext not in EXTS:
            print(f"⚠️ Unsupported: {path}")
            continue

        # extract text
        try:
            if ext == "docx":
                text = extract_text_from_docx(path)
            elif ext == "pdf":
                text = extract_text_from_pdf(path)
            elif ext == "pages":
                text = extract_text_from_pages(path)
            else:
                text = extract_text_from_txt(path)
        except Exception as e:
            print(f"❌ Extract error ({path}): {e}")
            continue

        # call Deepseek
        try:
            people = analyze_text_with_deepseek(text)
        except Exception as e:
            print(f"❌ Analyze error ({path}): {e}")
            continue

        # upsert each person *and move to end* so recency is preserved
        for person in people:
            name = person.get("name", "").strip()
            interests = person.get("interests", [])
            key = (os.path.basename(path), name)
            if key in od:
                # merge unique interests
                merged = list(set(od[key]).union(interests))
                # remove old position, re-insert at end
                del od[key]
                od[key] = merged
                print(f"🔄 Updated (moved to bottom): {key}")
            else:
                od[key] = interests
                print(f"➕ Added: {key}")

    # 3) Rewrite CSV in insertion (recency) order
    write_ordereddict_to_csv(args.out, od)
    print(f"\nDone — CSV now up-to-date at {args.out}")


if __name__ == "__main__":
    main()
