#!/usr/bin/env python3
"""
Invoice OCR Extractor - FAST VERSION (Multiprocessing)
- Paralel processing pakai semua CPU core
- OCR 1x per gambar (teks + layout sekaligus dengan psm 6)
- Resize cukup 1.5x (tradeoff kecepatan vs akurasi)
"""

import re
import os
import sys
import pytesseract
from PIL import Image
import pandas as pd
from pathlib import Path
from multiprocessing import Pool, cpu_count

# ─── CONFIG ────────────────────────────────────────────────────────────────────
IMAGE_DIR   = Path("batch1_1")
OUTPUT_XLSX = "batch1_1_extracted.xlsx"
OUTPUT_CSV  = "batch1_1_extracted.csv"
SCALE       = 1.5   # Resize factor (1.5x cukup untuk teks besar)
# ───────────────────────────────────────────────────────────────────────────────


def clean_num(s: str) -> str:
    s = str(s).strip().replace(" ", "")
    if re.search(r'\d{1,3}\.\d{3},\d{1,2}', s):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    elif "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    return s


def process_image(img_path: Path) -> list[dict]:
    """Proses satu gambar: OCR + parsing → list of row dicts."""
    try:
        img = Image.open(img_path)
        w, h = img.size
        img = img.resize((int(w * SCALE), int(h * SCALE)), Image.LANCZOS)

        # --- OCR dengan data posisi (satu kali saja) ---
        data = pytesseract.image_to_data(
            img, lang="eng",
            config="--psm 6",
            output_type=pytesseract.Output.DICT
        )

        img_width = int(w * SCALE)
        mid_x = img_width * 0.5

        # Susun kata-kata menjadi baris dengan info sisi (kiri/kanan)
        n = len(data["text"])
        line_map = {}   # (block,par,line) → {"left":[], "right":[], "all":[]}
        for i in range(n):
            conf = int(data["conf"][i])
            if conf < 0:
                continue
            txt = data["text"][i].strip()
            if not txt:
                continue
            key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
            if key not in line_map:
                line_map[key] = {"left": [], "right": [], "all": [], "top": data["top"][i]}
            side = "left" if data["left"][i] < mid_x else "right"
            line_map[key][side].append(txt)
            line_map[key]["all"].append(txt)

        # Sort baris berdasarkan posisi vertikal
        sorted_keys = sorted(line_map.keys(), key=lambda k: line_map[k]["top"])

        def get_line(key, side="all"):
            return " ".join(line_map[key][side])

        full_lines = [get_line(k) for k in sorted_keys]
        full_text  = "\n".join(full_lines)

        # ── Invoice No ──────────────────────────────────────────
        invoice_no = ""
        for l in full_lines:
            m = re.search(r"[Ii]nvoice\s+no[:\.]?\s*([0-9A-Za-z\-/]+)", l)
            if m:
                invoice_no = m.group(1).strip()
                break

        # ── Date of Issue ───────────────────────────────────────
        date_of_issue = ""
        for i, l in enumerate(full_lines):
            if re.search(r"[Dd]ate\s+of\s+issue", l):
                combined = " ".join(full_lines[i:i+3])
                m = re.search(r"(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", combined)
                if m:
                    date_of_issue = m.group(1)
                break

        # ── Seller & Client (pakai posisi kolom kiri/kanan) ─────
        seller_parts, client_parts = [], []
        collecting = False
        seller_header_top = -1

        for k in sorted_keys:
            lt = get_line(k, "left")
            rt = get_line(k, "right")
            top = line_map[k]["top"]

            if re.search(r"\bSeller\b", lt, re.IGNORECASE) and not collecting:
                collecting = True
                seller_header_top = top
                continue

            if collecting:
                if re.search(r"\bITEMS\b|No\.\s+Description", lt + rt, re.IGNORECASE):
                    break
                if lt and not re.match(r"(Tax\s*Id|IBAN)", lt, re.IGNORECASE):
                    seller_parts.append(lt)
                if rt and not re.match(r"(Tax\s*Id|IBAN|Client\s*:)", rt, re.IGNORECASE):
                    client_parts.append(rt)

        def get_name_loc(parts):
            name = parts[0].strip() if parts else ""
            loc  = ""
            for p in parts:
                if re.search(r",\s*[A-Z]{2}\s+\d{4,5}|[A-Z]{2}\s+\d{4,5}|\bAP\s+\d+|\bFPO\b|\bDPO\b", p):
                    loc = p.strip()
                    break
            if not loc and len(parts) > 1:
                loc = parts[-1].strip()
            return name, loc

        seller_name, seller_loc = get_name_loc(seller_parts)
        client_name, client_loc = get_name_loc(client_parts)

        # ── Items Table ─────────────────────────────────────────
        items_start = summary_idx = table_header = -1
        for i, l in enumerate(full_lines):
            if items_start < 0 and re.search(r"\bITEMS\b", l, re.IGNORECASE):
                items_start = i
            if re.search(r"\bSUMMARY\b", l, re.IGNORECASE):
                summary_idx = i
                break
        if summary_idx < 0:
            summary_idx = len(full_lines)

        search_from = items_start if items_start >= 0 else 0
        for i in range(search_from, min(search_from + 15, len(full_lines))):
            if re.search(r"\bNo\.\b.*\bDescription\b|\bDescription\b.*\bQty\b",
                         full_lines[i], re.IGNORECASE):
                table_header = i
                break

        rows = []
        if table_header >= 0:
            item_lines = full_lines[table_header + 1 : summary_idx]

            raw_items, cur = [], []
            for l in item_lines:
                if not l:
                    if cur: raw_items.append(" ".join(cur)); cur = []
                    continue
                # Deteksi awal item baru: "1." atau "1:" atau baris dengan QTY+Unit
                is_new_item = False
                if re.match(r"^\d+[\.:]", l):
                    is_new_item = True
                    l = re.sub(r"^(\d+):", r"\1.", l)
                elif re.search(r"\d[\d,\.]*\s+(?:each|pcs|pc|set|szt|ks|sth|kom|unit|ud|unt)\b", l, re.IGNORECASE):
                    is_new_item = True
                    
                if is_new_item:
                    if cur: raw_items.append(" ".join(cur))
                    cur = [l]
                elif cur:
                    cur.append(l)
            if cur: raw_items.append(" ".join(cur))

            for raw in raw_items:
                raw = raw.strip()
                if not raw: continue

                item_no = desc = qty = net_price = ""

                # Pattern A: dengan unit (each/pcs/...)
                m = re.search(
                    r"^(.*?)\s+(\d[\d,\.]*)\.?\s+(?:each|pcs|pc|set|szt|ks|sth|kom|unit|ud|unt)\b\s+([\d,\.]+)",
                    raw, re.IGNORECASE
                )
                if m:
                    desc = m.group(1).strip()
                    desc = re.sub(r"^(\d+[\.:]?|[\w]{1,3})\s+", "", desc)
                    qty = clean_num(m.group(2))
                    net_price = clean_num(m.group(3))

                # Pattern B: tanpa unit (fallback)
                if not desc:
                    m = re.match(
                        r"^(\d+)\.\s+(.+?)\s+"
                        r"(\d{1,4}[,\.]?\d{0,2})\s+"
                        r"([\d,\.]{2,})",
                        raw, re.IGNORECASE
                    )
                    if m:
                        desc = m.group(2).strip()
                        qty = clean_num(m.group(3))
                        net_price = clean_num(m.group(4))

                # Pattern C: fallback super basic
                if not desc:
                    nums = re.findall(r"\d[\d,\.]*", raw)
                    desc = re.sub(r"^[\w]{1,3}\s+", "", raw)
                    desc = re.sub(r"\s+\d[\d,\. ]+.*$", "", desc).strip()
                    qty = clean_num(nums[1]) if len(nums) > 1 else ""
                    net_price = clean_num(nums[2]) if len(nums) > 2 else ""

                rows.append({
                    "Nama File": img_path.name,
                    "Invoice No.": invoice_no,
                    "Date of issue :": date_of_issue,
                    "Seller": seller_name,
                    "Location Seller": seller_loc,
                    "Client": client_name,
                    "Location Client": client_loc,
                    "Item Description": desc,
                    "QTY": qty,
                    "Net Price": net_price,
                })

        if not rows:
            rows.append({
                "Nama File": img_path.name,
                "Invoice No.": invoice_no,
                "Date of issue :": date_of_issue,
                "Seller": seller_name,
                "Location Seller": seller_loc,
                "Client": client_name,
                "Location Client": client_loc,
                "Item Description": "", "QTY": "", "Net Price": "",
            })

        return rows

    except Exception as e:
        return [{
            "Nama File": img_path.name,
            "Invoice No.": f"ERROR: {e}",
            "Date of issue :": "", "Seller": "", "Location Seller": "",
            "Client": "", "Location Client": "",
            "Item Description": "", "QTY": "", "Net Price": "",
        }]


def worker(img_path_str: str) -> list[dict]:
    """Wrapper untuk multiprocessing (harus top-level, argumen string)."""
    return process_image(Path(img_path_str))


def main():
    image_files = sorted(IMAGE_DIR.glob("*.jpg")) + sorted(IMAGE_DIR.glob("*.png"))
    total = len(image_files)

    cores = cpu_count()
    workers = max(1, cores - 1)  # sisakan 1 core untuk OS
    print(f"📂 {total} gambar ditemukan | 🖥️  {cores} CPU core → pakai {workers} worker")
    print(f"⚡ Mulai ekstraksi paralel...\n")

    all_rows = []
    done = 0

    with Pool(processes=workers) as pool:
        for rows in pool.imap_unordered(worker, [str(p) for p in image_files], chunksize=4):
            all_rows.extend(rows)
            done += 1
            if done % 50 == 0 or done == total:
                print(f"  ✔ [{done}/{total}] selesai diproses")

    # Sort berdasarkan filename agar urut
    all_rows.sort(key=lambda r: r["Nama File"])

    COLUMNS = ["Nama File", "Invoice No.", "Date of issue :",
               "Seller", "Location Seller",
               "Client", "Location Client",
               "Item Description", "QTY", "Net Price"]

    df = pd.DataFrame(all_rows, columns=COLUMNS)

    # ── Excel ──────────────────────────────────────────────────
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Invoices")
        ws = writer.sheets["Invoices"]
        col_widths = {"Nama File":20, "Invoice No.":15, "Date of issue :":14,
                      "Seller":30, "Location Seller":30,
                      "Client":30, "Location Client":30,
                      "Item Description":50, "QTY":8, "Net Price":12}
        for col_cells in ws.columns:
            header = col_cells[0].value
            ws.column_dimensions[col_cells[0].column_letter].width = col_widths.get(header, 15)

    # ── CSV ────────────────────────────────────────────────────
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    errors = [r for r in all_rows if str(r.get("Invoice No.", "")).startswith("ERROR")]
    print(f"\n✅ SELESAI! {len(all_rows)} baris dari {total} invoice.")
    print(f"   📊 Excel : {OUTPUT_XLSX}")
    print(f"   📄 CSV   : {OUTPUT_CSV}")
    if errors:
        print(f"   ⚠️  {len(errors)} file error.")


if __name__ == "__main__":
    main()
