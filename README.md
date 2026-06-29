# invoice-ocr-extractor
Fast invoice OCR extractor using Tesseract + regex + multiprocessing
# 🧾 Invoice OCR Extractor

Fast, parallel invoice data extractor using Tesseract OCR + regex parsing.
Extracts structured data (invoice no, date, seller, client, items) from scanned invoices into Excel/CSV.

## ✨ Features

- ⚡ **Multiprocessing** — uses all CPU cores for batch processing
- 📐 **Layout-aware parsing** — auto-detects left/right columns (Seller vs Client)
- 🔄 **3-layer regex fallback** — robust item table extraction
- 🌍 **Number format agnostic** — handles EU (1.234,56) & US (1,234.56)
- 📊 **Dual output** — Excel + CSV

## 🛠️ Tech Stack

- Python 3.10+
- Tesseract OCR (`pytesseract`)
- Pillow (image preprocessing)
- Pandas + openpyxl (data export)
- Multiprocessing (parallelism)

## 🚀 Installation

```bash
#1. Install Tesseract OCR (system-level)
# Ubuntu/Debian:
sudo apt install tesseract-ocr tesseract-ocr-ind

# macOS:
brew install tesseract tesseract-lang

