# AI Name & Interest Extractor

A Python utility to ingest `.docx`, `.pdf`, `.pages`, or `.txt` files, call the Deepseek API to extract every person’s name and interests, and maintain a recency-ordered CSV of results.

## Features

- Supports Microsoft Word (`.docx`), PDF, Apple Pages (`.pages`), and plain text (`.txt`).
- Unzips `.pages` to grab the embedded PDF preview.
- Sends your text to Deepseek and parses a JSON array of `{ name, interests }`.
- Upserts into `results.csv`, merging interests and re-ordering rows by last update.
