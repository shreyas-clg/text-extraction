#!/usr/bin/env python3
"""Extract invoice fields from PDF(s) into an Excel workbook.

Usage:
    python extract.py invoice1.pdf invoice2.pdf ...
    python extract.py *.pdf -o result.xlsx

Handles two invoice layouts, auto-detected from page 1:
    TYPE A - DISTRIBUTION FRANPRIX
    TYPE B - SEDIFRAIS

PDFs are digital (have a text layer), so text is extracted directly with
pdfplumber - no OCR engine needed.
"""
import argparse
import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook

# Space-like separators used as thousands grouping: normal, non-breaking ( ),
# narrow non-breaking ( ).
_SP = "[   ]"
# A French amount: thousands in blocks of 3, comma, 2 decimals. Strict grouping
# stops two space-separated amounts being merged. e.g. "7 842,25", "342,32".
NUM = rf"\d{{1,3}}(?:{_SP}\d{{3}})*,\d{{2}}"
DATE = r"\d{2}/\d{2}/\d{2,4}"


def find(pattern, text, group=1, flags=0):
    """Return captured group of first match, or '' if none."""
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else ""


def clean_num(s):
    """Normalize a French amount to a plain string '7842.25' (or '' if empty)."""
    if not s:
        return ""
    for sp in (" ", " ", " "):
        s = s.replace(sp, "")
    return s.replace(",", ".")


# ---- Field columns per type (order = spreadsheet column order) ----
COLUMNS_A = [
    "Fichier", "N_FACTURE", "Date", "N_CLIENT", "PRELEVEMENT_ECHEANCE",
    "TOTAUX_Mts_PUBLIC_TTC", "TOTAUX_Mts_MARCH_HTVA", "TOTAUX_Mts_TVA",
    "MONTANT_A_PAYER",
]
COLUMNS_B = [
    "Fichier", "Livre_a", "FACTURE_N", "Date_facture", "Date_echeance",
    "Montant_brut_hors_tva", "Montant_TVA", "Montant_a_payer",
]


def parse_type_a(pages_text, filename):
    """DISTRIBUTION FRANPRIX."""
    # Header row on page 1: FACTURE(6d) DATE(dd/mm/yy) CLIENT(6d)
    hdr = re.search(rf"(\d{{6}})\s+({DATE})\s+(\d{{6}})", pages_text[0])
    n_facture = hdr.group(1) if hdr else ""
    date = hdr.group(2) if hdr else ""
    n_client = hdr.group(3) if hdr else ""

    # Totals live on the last page.
    last = pages_text[-1]
    echeance = find(rf"PRELEVEMENT ECHEANCE\s+({DATE})", last)
    tot = re.search(rf"TOTAUX\s+({NUM})\s+({NUM})\s+({NUM})", last)
    public_ttc = clean_num(tot.group(1)) if tot else ""
    march_htva = clean_num(tot.group(2)) if tot else ""
    tva = clean_num(tot.group(3)) if tot else ""
    # MONTANT A PAYER = amount immediately before "POIDS LIVRE".
    montant = clean_num(find(rf"({NUM}){_SP}*POIDS LIVRE", last))

    return {
        "Fichier": filename, "N_FACTURE": n_facture, "Date": date,
        "N_CLIENT": n_client, "PRELEVEMENT_ECHEANCE": echeance,
        "TOTAUX_Mts_PUBLIC_TTC": public_ttc, "TOTAUX_Mts_MARCH_HTVA": march_htva,
        "TOTAUX_Mts_TVA": tva, "MONTANT_A_PAYER": montant,
    }


def parse_type_b(pages_text, filename):
    """SEDIFRAIS."""
    p1 = pages_text[0]
    livre_a = find(r"Livré à\s*:\s*(\d+)", p1)
    facture_n = find(r"N°\s*(\d+)", p1)
    date_facture = find(rf"Date facture\s*:\s*({DATE})", p1)
    date_echeance = find(rf"Date échéance\s*:\s*({DATE})", p1)

    # Totals block (last page).
    last = pages_text[-1]
    brut = clean_num(find(rf"Montant brut hors tva\s+({NUM})", last))
    tva = clean_num(find(rf"Montant TVA\s+({NUM})", last))
    a_payer = clean_num(find(rf"Montant à payer\s+({NUM})", last))

    return {
        "Fichier": filename, "Livre_a": livre_a, "FACTURE_N": facture_n,
        "Date_facture": date_facture, "Date_echeance": date_echeance,
        "Montant_brut_hors_tva": brut, "Montant_TVA": tva,
        "Montant_a_payer": a_payer,
    }


def detect_type(page1_text):
    """Return 'A', 'B', or None. Decide on page 1 only - the CGV pages of a
    SEDIFRAIS invoice also mention FRANPRIX, so whole-doc matching is unsafe."""
    if "DISTRIBUTION FRANPRIX" in page1_text:
        return "A"
    if "SEDIFRAIS" in page1_text:
        return "B"
    return None


def process(path):
    """Extract one PDF. Returns (type, row_dict) or (None, None) if unknown."""
    with pdfplumber.open(path) as pdf:
        pages_text = [(pg.extract_text() or "") for pg in pdf.pages]
    kind = detect_type(pages_text[0] if pages_text else "")
    name = Path(path).name
    if kind == "A":
        return "A", parse_type_a(pages_text, name)
    if kind == "B":
        return "B", parse_type_b(pages_text, name)
    return None, None


def write_xlsx(rows_a, rows_b, out_path):
    wb = Workbook()
    wb.remove(wb.active)  # drop default sheet
    for title, cols, rows in [("TypeA", COLUMNS_A, rows_a),
                              ("TypeB", COLUMNS_B, rows_b)]:
        ws = wb.create_sheet(title)
        ws.append(cols)
        for r in rows:
            ws.append([r.get(c, "") for c in cols])
    wb.save(out_path)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Extract invoice fields to Excel.")
    ap.add_argument("pdfs", nargs="+", help="PDF file(s)")
    ap.add_argument("-o", "--output", default="output.xlsx", help="output .xlsx")
    args = ap.parse_args(argv)

    rows_a, rows_b = [], []
    for p in args.pdfs:
        if not Path(p).is_file():
            print(f"SKIP (not found): {p}", file=sys.stderr)
            continue
        kind, row = process(p)
        if kind == "A":
            rows_a.append(row)
            print(f"[A] {Path(p).name}: {row}")
        elif kind == "B":
            rows_b.append(row)
            print(f"[B] {Path(p).name}: {row}")
        else:
            print(f"SKIP (unknown type): {p}", file=sys.stderr)

    write_xlsx(rows_a, rows_b, args.output)
    print(f"\nWrote {len(rows_a)} TypeA + {len(rows_b)} TypeB row(s) -> {args.output}")


if __name__ == "__main__":
    main()
