#!/usr/bin/env python3
import os
import zipfile
import tempfile
import json
import sqlite3
import PyPDF2
import requests
from docx import Document
import argparse

# ——— Configuration ———
DEEPSEEK_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_URL = "https://api.deepseek.ai/v1/analyze"

DB_SCHEMA = """
CREATE TABLE IF NOT EXISTS people (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    interests TEXT NOT NULL
);
"""


def init_db(db_path: str):
    conn = sqlite3.connect(db_path)
    conn.execute(DB_SCHEMA)
    conn.commit()
    return conn


def save_person(conn, name: str, interests: list):
    interests_json = json.dumps(interests, ensure_ascii=False)
    conn.execute(
        "INSERT INTO people (name, interests) VALUES (?, ?)",
        (name, interests_json),
    )
    conn.commit()


def extract_text_from_pdf(path: str) -> str:
    reader = PyPDF2.PdfReader(path)
    return "\n\n".join(page.extract_text() or "" for page in reader.pages)


def extract_text_from_pages(path: str) -> str:
    with zipfile.ZipFile(path, "r") as z:
        # find the embedded PDF preview
        candidates = [
            f for f in z.namelist() if f.startswith("QuickLook/") and f.endswith(".pdf")
        ]
        if not candidates:
            raise FileNotFoundError("No QuickLook/Preview.pdf found in the .pages file")
        with z.open(candidates[0]) as embedded_pdf, tempfile.NamedTemporaryFile(
            delete=False, suffix=".pdf"
        ) as tf:
            tf.write(embedded_pdf.read())
            return extract_text_from_pdf(tf.name)


def extract_text_from_docx(path: str) -> str:
    doc = Document(path)
    return "\n\n".join(p.text for p in doc.paragraphs if p.text)


def analyze_text_with_deepseek(text: str) -> dict:
    payload = {
        "prompt": (
            "Extract the person’s full name and their interests from the following text:\n\n"
            f"```\n{text}\n```\n\n"
            "Return JSON with keys `name` (string) and `interests` (list of strings)."
        ),
        "options": {"format": "json"},
    }
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_KEY}",
        "Content-Type": "application/json",
    }
    r = requests.post(DEEPSEEK_URL, json=payload, headers=headers)
    r.raise_for_status()
    return r.json()


def main():
    p = argparse.ArgumentParser(
        description="Extract name & interests from .pages/.pdf/.docx and store in SQLite via Deepseek"
    )
    p.add_argument(
        "input_files", nargs="+", help=".pages, .pdf or .docx files to process"
    )
    p.add_argument(
        "--db", default="people.db", help="SQLite database path (default: people.db)"
    )
    args = p.parse_args()

    conn = init_db(args.db)

    for path in args.input_files:
        if not os.path.isfile(path):
            print(f"⚠️ Skipping missing file: {path}")
            continue

        ext = path.lower().rsplit(".", 1)[-1]
        try:
            if ext == "pages":
                text = extract_text_from_pages(path)
            elif ext == "pdf":
                text = extract_text_from_pdf(path)
            elif ext == "docx":
                text = extract_text_from_docx(path)
            else:
                print(f"⚠️ Unsupported file type, skipping: {path}")
                continue
        except Exception as e:
            print(f"Error extracting text from {path}: {e}")
            continue

        try:
            data = analyze_text_with_deepseek(text)
            name = data.get("name", "").strip()
            interests = data.get("interests", [])
            if name and isinstance(interests, list):
                save_person(conn, name, interests)
                print(f"✔️ Saved: {name} → {interests}")
            else:
                print(f"❓ Unexpected response for {path}: {data}")
        except Exception as e:
            print(f"Error analyzing {path}: {e}")

    conn.close()


if __name__ == "__main__":
    main()
