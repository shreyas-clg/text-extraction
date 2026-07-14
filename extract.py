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
from datetime import datetime
import glob
import logging
import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook

logger = logging.getLogger(__name__)

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

    if not n_facture:
        logger.warning("[%s] Field 'N_FACTURE' could not be parsed", filename)
    if not date:
        logger.warning("[%s] Field 'Date' could not be parsed", filename)
    if not n_client:
        logger.warning("[%s] Field 'N_CLIENT' could not be parsed", filename)

    # Totals live on the last page.
    last = pages_text[-1]
    echeance = find(rf"PRELEVEMENT ECHEANCE\s+({DATE})", last)
    tot = re.search(rf"TOTAUX\s+({NUM})\s+({NUM})\s+({NUM})", last)
    public_ttc = clean_num(tot.group(1)) if tot else ""
    march_htva = clean_num(tot.group(2)) if tot else ""
    tva = clean_num(tot.group(3)) if tot else ""
    # MONTANT A PAYER = amount immediately before "POIDS LIVRE".
    montant = clean_num(find(rf"({NUM}){_SP}*POIDS LIVRE", last))

    if not montant:
        logger.warning("[%s] Field 'MONTANT_A_PAYER' could not be parsed", filename)

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

    if not facture_n:
        logger.warning("[%s] Field 'FACTURE_N' could not be parsed", filename)
    if not date_facture:
        logger.warning("[%s] Field 'Date_facture' could not be parsed", filename)

    # Totals block (last page).
    last = pages_text[-1]
    brut = clean_num(find(rf"Montant brut hors tva\s+({NUM})", last))
    tva = clean_num(find(rf"Montant TVA\s+({NUM})", last))
    a_payer = clean_num(find(rf"Montant à payer\s+({NUM})", last))

    if not a_payer:
        logger.warning("[%s] Field 'Montant_a_payer' could not be parsed", filename)

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
    name = Path(path).name
    pages_text = []
    
    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            logger.warning("[%s] PDF contains no pages", name)
        for idx, pg in enumerate(pdf.pages):
            try:
                text = pg.extract_text() or ""
                pages_text.append(text)
            except Exception as e:
                logger.error("[%s] Failed to extract text from page %d: %s", name, idx + 1, e)
                pages_text.append("")

    if not pages_text or not any(pages_text):
        logger.warning("[%s] No text extracted from any page of the PDF", name)
        return None, None

    kind = detect_type(pages_text[0])
    if kind == "A":
        return "A", parse_type_a(pages_text, name)
    if kind == "B":
        return "B", parse_type_b(pages_text, name)
    
    logger.warning("[%s] Unknown invoice layout. Could not determine type A or B", name)
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


def discover_files(inputs, batch_folders=None, recursive=False):
    """Scan and resolve all unique PDF files based on inputs and batch folders."""
    files = []
    seen = set()

    def add_file(path):
        try:
            resolved = Path(path).resolve()
            if resolved not in seen:
                seen.add(resolved)
                files.append(Path(path))
        except Exception:
            if path not in seen:
                seen.add(path)
                files.append(Path(path))

    def scan_dir(dir_path):
        pattern = "*.[pP][dD][fF]"
        if recursive:
            for p in Path(dir_path).rglob(pattern):
                if p.is_file():
                    add_file(p)
        else:
            for p in Path(dir_path).glob(pattern):
                if p.is_file():
                    add_file(p)

    targets = []
    if inputs:
        targets.extend(inputs)
    if batch_folders:
        targets.extend(batch_folders)

    for item in targets:
        item_str = str(item)
        if any(char in item_str for char in "*?[]"):
            matched = glob.glob(item_str, recursive=recursive)
            for m in matched:
                p = Path(m)
                if p.is_dir():
                    scan_dir(p)
                elif p.is_file():
                    if p.suffix.lower() == ".pdf":
                        add_file(p)
            continue

        p = Path(item)
        if p.is_dir():
            scan_dir(p)
        elif p.is_file():
            if p.suffix.lower() == ".pdf":
                add_file(p)
            else:
                logger.warning("Skipped: %s is not a PDF file", p)
        else:
            logger.warning("Skipped: %s not found", p)

    return files


def main(argv=None):
    ap = argparse.ArgumentParser(description="Extract invoice fields to Excel.")
    ap.add_argument("inputs", nargs="*", help="PDF file(s), folder(s), or glob pattern(s)")
    ap.add_argument("-b", "--batch", "--folder", nargs="+", dest="batch_folders", help="Folder(s) to scan for PDF files automatically")
    ap.add_argument("-r", "--recursive", action="store_true", help="Scan folders recursively")
    ap.add_argument("-o", "--output", default="output.xlsx", help="output .xlsx")
    args = ap.parse_args(argv)

    if not args.inputs and not args.batch_folders:
        ap.error("No input files or folders specified. Please provide at least one PDF file/glob or use --batch/--folder.")

    # --- Setup Logging ---
    output_path = Path(args.output)
    log_dir = output_path.parent
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        log_dir = Path(".")
        
    date_str = datetime.now().strftime("%Y%m%d")
    log_path = log_dir / f"extraction_log_{date_str}.log"

    log_format = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_path, encoding="utf-8")
        ]
    )

    pdf_paths = discover_files(args.inputs, args.batch_folders, args.recursive)
    if not pdf_paths:
        logger.error("No PDF files were discovered for processing.")
        sys.exit(1)

    rows_a, rows_b = [], []
    for p in pdf_paths:
        try:
            kind, row = process(p)
            if kind == "A":
                rows_a.append(row)
                logger.info("Processed Type A invoice: %s -> %s", p.name, row)
            elif kind == "B":
                rows_b.append(row)
                logger.info("Processed Type B invoice: %s -> %s", p.name, row)
        except Exception as e:
            logger.error("Failed to process invoice %s: %s", p.name, e, exc_info=True)

    write_xlsx(rows_a, rows_b, args.output)
    logger.info("Wrote %d TypeA + %d TypeB row(s) -> %s", len(rows_a), len(rows_b), args.output)


if __name__ == "__main__":
    main()
