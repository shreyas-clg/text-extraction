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
from typing import Dict, Any, List, Optional, Tuple, Set

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


# ---- Common Utilities ----

def find(pattern: str, text: str, group: int = 1, flags: re.RegexFlag = re.IGNORECASE) -> str:
    """Return captured group of first match, or '' if none."""
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else ""


def clean_num(s: str) -> str:
    """Normalize a French amount to a plain string '7842.25' (or '' if empty)."""
    if not s:
        return ""
    for sp in (" ", " ", " "):
        s = s.replace(sp, "")
    return s.replace(",", ".")


# ---- Base Parser and Registry ----

class BaseParser:
    """Base class for all invoice parsers."""

    def detect(self, page1_text: str) -> bool:
        """Return True if this parser should handle the given invoice."""
        raise NotImplementedError

    def parse(self, pages_text: List[str], filename: str) -> Dict[str, Any]:
        """Extract fields and return dictionary of properties."""
        raise NotImplementedError

    @property
    def columns(self) -> List[str]:
        """List of output Excel columns for this invoice type."""
        raise NotImplementedError

    @property
    def sheet_name(self) -> str:
        """Name of the Excel sheet to save this invoice type's results."""
        raise NotImplementedError


class ParserRegistry:
    """Registry to keep track of invoice parsers."""

    def __init__(self) -> None:
        self._parsers: Dict[str, BaseParser] = {}

    def register(self, name: str, parser: BaseParser) -> None:
        """Register a new invoice parser."""
        self._parsers[name] = parser

    def get_all(self) -> Dict[str, BaseParser]:
        """Get all registered parsers."""
        return self._parsers

    def detect_and_process(self, pages_text: List[str], filename: str) -> Tuple[Optional[str], Optional[Dict[str, Any]], Optional[BaseParser]]:
        """Find the matching parser, run parse, and return (name, row_dict, parser)."""
        if not pages_text:
            return None, None, None
        for name, parser in self._parsers.items():
            if parser.detect(pages_text[0]):
                return name, parser.parse(pages_text, filename), parser
        return None, None, None


# Global registry instance
registry = ParserRegistry()


def register_parser(name: str, parser: BaseParser) -> None:
    """API function to register new invoice types."""
    registry.register(name, parser)


# ---- Concrete Parsers ----

class TypeAParser(BaseParser):
    """DISTRIBUTION FRANPRIX / Grid layout parser."""

    @property
    def sheet_name(self) -> str:
        return "TypeA"

    @property
    def columns(self) -> List[str]:
        return [
            "Fichier", "N_FACTURE", "Date", "N_CLIENT", "PRELEVEMENT_ECHEANCE",
            "TOTAUX_Mts_PUBLIC_TTC", "TOTAUX_Mts_MARCH_HTVA", "TOTAUX_Mts_TVA",
            "MONTANT_A_PAYER", "Validation_Status", "Notes",
        ]

    def detect(self, page1_text: str) -> bool:
        return "FACTURE REFERENCE INTERNE" in page1_text or "Mts PUBLIC" in page1_text

    def parse(self, pages_text: List[str], filename: str) -> Dict[str, Any]:
        # Header row on page 1: FACTURE(4-10d) DATE(dd/mm/yy) CLIENT(4-10d)
        hdr = re.search(rf"(\d{{4,10}})\s+({DATE})\s+(\d{{4,10}})", pages_text[0])
        n_facture = hdr.group(1) if hdr else ""
        date = hdr.group(2) if hdr else ""
        n_client = hdr.group(3) if hdr else ""

        if not n_facture:
            logger.warning("[%s] Field 'N_FACTURE' could not be parsed", filename)
        if not date:
            logger.warning("[%s] Field 'Date' could not be parsed", filename)
        if not n_client:
            logger.warning("[%s] Field 'N_CLIENT' could not be parsed", filename)

        echeance = ""
        for txt in pages_text:
            echeance = find(rf"PRELEVEMENT ECHEANCE\s+({DATE})", txt)
            if echeance:
                break

        # Search all pages for totals.
        totals_page_text = ""
        for txt in pages_text:
            if "TOTAUX" in txt or "POIDS LIVRE" in txt:
                totals_page_text = txt
                break
        if not totals_page_text:
            totals_page_text = pages_text[-1]

        tot_line = find(r"(TOTAUX\s+.*)", totals_page_text)
        tot_vals = re.findall(NUM, tot_line) if tot_line else []
        
        if len(tot_vals) >= 3:
            public_ttc = clean_num(tot_vals[0])
            march_htva = clean_num(tot_vals[1])
            tva = clean_num(tot_vals[2])
        elif len(tot_vals) == 2:
            public_ttc = ""
            march_htva = clean_num(tot_vals[0])
            tva = clean_num(tot_vals[1])
        elif len(tot_vals) == 1:
            public_ttc = ""
            march_htva = clean_num(tot_vals[0])
            tva = ""
        else:
            public_ttc = ""
            march_htva = ""
            tva = ""

        montant = clean_num(find(rf"({NUM}){_SP}*POIDS LIVRE", totals_page_text))
        if not montant:
            # Fallback to finding MONTANT A PAYER / P.A.Y.E.R
            m = re.search(r"MONTANT\s+A\s*(?:P\.A\.Y\.E\.R|PAYER)", totals_page_text, re.IGNORECASE)
            if m:
                start_pos = m.end()
                m_num = re.search(rf"({NUM})", totals_page_text[start_pos:])
                if m_num:
                    montant = clean_num(m_num.group(1))

        if not montant:
            logger.warning("[%s] Field 'MONTANT_A_PAYER' could not be parsed", filename)

        return {
            "Fichier": filename,
            "N_FACTURE": n_facture,
            "Date": date,
            "N_CLIENT": n_client,
            "PRELEVEMENT_ECHEANCE": echeance,
            "TOTAUX_Mts_PUBLIC_TTC": public_ttc,
            "TOTAUX_Mts_MARCH_HTVA": march_htva,
            "TOTAUX_Mts_TVA": tva,
            "MONTANT_A_PAYER": montant,
        }


class TypeBParser(BaseParser):
    """SEDIFRAIS / List layout parser."""

    @property
    def sheet_name(self) -> str:
        return "TypeB"

    @property
    def columns(self) -> List[str]:
        return [
            "Fichier", "Livre_a", "FACTURE_N", "Date_facture", "Date_echeance",
            "Montant_brut_hors_tva", "Montant_TVA", "Montant_a_payer",
            "Validation_Status", "Notes",
        ]

    def detect(self, page1_text: str) -> bool:
        return "Montant brut hors tva" in page1_text or "Date facture" in page1_text or "Livré à" in page1_text

    def parse(self, pages_text: List[str], filename: str) -> Dict[str, Any]:
        p1 = pages_text[0]
        livre_a = find(r"Livré à\s*:?\s*(\d+)", p1)
        facture_n = find(r"N°\s*:?\s*(\d+)", p1)
        date_facture = find(rf"Date\s+facture\s*:?\s*({DATE})", p1)
        date_echeance = find(rf"Date\s+échéance\s*:?\s*({DATE})", p1)

        if not facture_n:
            logger.warning("[%s] Field 'FACTURE_N' could not be parsed", filename)
        if not date_facture:
            logger.warning("[%s] Field 'Date_facture' could not be parsed", filename)

        # Search all pages for totals.
        totals_page_text = ""
        for txt in pages_text:
            if "Montant à payer" in txt or "Montant brut hors tva" in txt:
                totals_page_text = txt
                break
        if not totals_page_text:
            totals_page_text = pages_text[-1]

        brut = clean_num(find(rf"Montant brut hors tva\s+({NUM})", totals_page_text))
        tva = clean_num(find(rf"Montant TVA\s+({NUM})", totals_page_text))
        a_payer = clean_num(find(rf"Montant à payer\s+({NUM})", totals_page_text))

        if not a_payer:
            logger.warning("[%s] Field 'Montant_a_payer' could not be parsed", filename)

        return {
            "Fichier": filename,
            "Livre_a": livre_a,
            "FACTURE_N": facture_n,
            "Date_facture": date_facture,
            "Date_echeance": date_echeance,
            "Montant_brut_hors_tva": brut,
            "Montant_TVA": tva,
            "Montant_a_payer": a_payer,
        }


# Register defaults
registry.register("TypeA", TypeAParser())
registry.register("TypeB", TypeBParser())


# ---- Legacy / Compatibility Helpers ----

def detect_type(page1_text: str) -> Optional[str]:
    """Helper for backward compatibility."""
    if registry.get_all()["TypeA"].detect(page1_text):
        return "A"
    if registry.get_all()["TypeB"].detect(page1_text):
        return "B"
    return None


def parse_type_a(pages_text: List[str], filename: str) -> Dict[str, Any]:
    """Helper for backward compatibility."""
    return registry.get_all()["TypeA"].parse(pages_text, filename)


def parse_type_b(pages_text: List[str], filename: str) -> Dict[str, Any]:
    """Helper for backward compatibility."""
    return registry.get_all()["TypeB"].parse(pages_text, filename)


# ---- Main Pipeline Functions ----

def process(path: Path) -> Tuple[Optional[str], Optional[Dict[str, Any]]]:
    """Extract one PDF. Returns (type_name, row_dict) or (None, None) if unknown."""
    name = Path(path).name
    pages_text: List[str] = []
    
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

    kind_name, row, _ = registry.detect_and_process(pages_text, name)
    if kind_name and row:
        return kind_name, row
    
    logger.warning("[%s] Unknown invoice layout. Could not determine type A or B", name)
    return None, None


def write_xlsx(results: Dict[str, List[Dict[str, Any]]], out_path: str) -> None:
    wb = Workbook()
    wb.remove(wb.active)  # drop default sheet
    for name, parser in registry.get_all().items():
        ws = wb.create_sheet(parser.sheet_name)
        ws.append(parser.columns)
        rows = results.get(name, [])
        for r in rows:
            ws.append([r.get(c, "") for c in parser.columns])
    wb.save(out_path)


def discover_files(inputs: List[str], batch_folders: Optional[List[str]] = None, recursive: bool = False) -> List[Path]:
    """Scan and resolve all unique PDF files based on inputs and batch folders."""
    files: List[Path] = []
    seen: Set[Path] = set()

    def add_file(path: Path) -> None:
        try:
            resolved = Path(path).resolve()
            if resolved not in seen:
                seen.add(resolved)
                files.append(Path(path))
        except Exception:
            if path not in seen:
                seen.add(path)
                files.append(Path(path))

    def scan_dir(dir_path: Path) -> None:
        pattern = "*.[pP][dD][fF]"
        if recursive:
            for p in Path(dir_path).rglob(pattern):
                if p.is_file():
                    add_file(p)
        else:
            for p in Path(dir_path).glob(pattern):
                if p.is_file():
                    add_file(p)

    targets: List[Any] = []
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


def validate_invoice(kind: str, row: Dict[str, Any]) -> Tuple[str, str]:
    """
    Validate the parsed row data.
    Returns (status, notes) where status is 'VALID' or 'SUSPICIOUS' and notes is a string.
    """
    status = "VALID"
    notes: List[str] = []

    def is_valid_date(d_str: str) -> bool:
        if not d_str:
            return False
        return bool(re.match(r"^\d{2}/\d{2}/\d{2,4}$", d_str))

    def is_valid_amount(a_str: str) -> bool:
        if not a_str:
            return False
        try:
            val = float(a_str)
            return val >= 0
        except ValueError:
            return False

    if kind == "TypeA":
        req_fields = ["N_FACTURE", "Date", "N_CLIENT", "MONTANT_A_PAYER"]
        for f in req_fields:
            if not row.get(f):
                status = "SUSPICIOUS"
                notes.append(f"Missing required field {f}")

        d = row.get("Date")
        if d and not is_valid_date(d):
            status = "SUSPICIOUS"
            notes.append(f"Invalid date format: {d}")

        echeance = row.get("PRELEVEMENT_ECHEANCE")
        if echeance and not is_valid_date(echeance):
            status = "SUSPICIOUS"
            notes.append(f"Invalid echeance date format: {echeance}")

        amt_fields = ["TOTAUX_Mts_PUBLIC_TTC", "TOTAUX_Mts_MARCH_HTVA", "TOTAUX_Mts_TVA", "MONTANT_A_PAYER"]
        for f in amt_fields:
            val = row.get(f)
            if val and not is_valid_amount(val):
                status = "SUSPICIOUS"
                notes.append(f"Invalid amount format in {f}: {val}")

        pay_str = row.get("MONTANT_A_PAYER")
        if pay_str:
            try:
                pay_val = float(pay_str)
                if pay_val > 1000000:
                    status = "SUSPICIOUS"
                    notes.append(f"Suspiciously high amount: {pay_val}")
            except ValueError:
                pass

        htva_str = row.get("TOTAUX_Mts_MARCH_HTVA")
        tva_str = row.get("TOTAUX_Mts_TVA")
        if htva_str and tva_str and pay_str:
            try:
                htva = float(htva_str)
                tva = float(tva_str)
                pay = float(pay_str)
                if abs((htva + tva) - pay) > 5.00:
                    notes.append(f"Math check: HTVA ({htva}) + TVA ({tva}) != PAYER ({pay})")
            except ValueError:
                pass

    elif kind == "TypeB":
        req_fields = ["FACTURE_N", "Date_facture", "Montant_a_payer"]
        for f in req_fields:
            if not row.get(f):
                status = "SUSPICIOUS"
                notes.append(f"Missing required field {f}")

        d = row.get("Date_facture")
        if d and not is_valid_date(d):
            status = "SUSPICIOUS"
            notes.append(f"Invalid date format: {d}")

        echeance = row.get("Date_echeance")
        if echeance and not is_valid_date(echeance):
            status = "SUSPICIOUS"
            notes.append(f"Invalid echeance date format: {echeance}")

        amt_fields = ["Montant_brut_hors_tva", "Montant_TVA", "Montant_a_payer"]
        for f in amt_fields:
            val = row.get(f)
            if val and not is_valid_amount(val):
                status = "SUSPICIOUS"
                notes.append(f"Invalid amount format in {f}: {val}")

        pay_str = row.get("Montant_a_payer")
        if pay_str:
            try:
                pay_val = float(pay_str)
                if pay_val > 1000000:
                    status = "SUSPICIOUS"
                    notes.append(f"Suspiciously high amount: {pay_val}")
            except ValueError:
                pass

        ht_str = row.get("Montant_brut_hors_tva")
        tva_str = row.get("Montant_TVA")
        if ht_str and tva_str and pay_str:
            try:
                ht = float(ht_str)
                tva = float(tva_str)
                pay = float(pay_str)
                if abs((ht + tva) - pay) > 5.00:
                    notes.append(f"Math check: HT ({ht}) + TVA ({tva}) != PAYER ({pay})")
            except ValueError:
                pass

    notes_str = "; ".join(notes)
    return status, notes_str


def main(argv: Optional[List[str]] = None) -> None:
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

    results: Dict[str, List[Dict[str, Any]]] = {name: [] for name in registry.get_all().keys()}
    for p in pdf_paths:
        try:
            kind_name, row = process(p)
            if kind_name and row:
                status, notes = validate_invoice(kind_name, row)
                row["Validation_Status"] = status
                row["Notes"] = notes
                results[kind_name].append(row)
                if status == "SUSPICIOUS":
                    logger.warning("Processed %s invoice (SUSPICIOUS): %s -> %s. Notes: %s", kind_name, p.name, row, notes)
                else:
                    logger.info("Processed %s invoice: %s -> %s", kind_name, p.name, row)
        except Exception as e:
            logger.error("Failed to process invoice %s: %s", p.name, e, exc_info=True)

    write_xlsx(results, args.output)
    
    total_written_msg = " + ".join(f"{len(rows)} {name}" for name, rows in results.items())
    logger.info("Wrote %s row(s) -> %s", total_written_msg, args.output)


if __name__ == "__main__":
    main()
