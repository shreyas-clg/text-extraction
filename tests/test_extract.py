import re
import pytest
from pathlib import Path
from extract import (
    find,
    clean_num,
    detect_type,
    validate_invoice,
    discover_files,
    registry,
    BaseParser,
    register_parser,
    process,
)

# Test helper functions
def test_clean_num():
    assert clean_num("7 842,25") == "7842.25"
    assert clean_num("342,32") == "342.32"
    assert clean_num("1 000,50") == "1000.50"
    assert clean_num("") == ""
    assert clean_num(None) == ""


def test_find():
    sample_text = "Date facture : 08/06/2026\nN° 970584"
    assert find(r"N°\s*(\d+)", sample_text) == "970584"
    assert find(r"Date facture\s*:\s*(\d{2}/\d{2}/\d{4})", sample_text) == "08/06/2026"
    assert find(r"Nonexistent", sample_text) == ""


# Test layout detection
def test_detect_type():
    txt_a = "Some text... FACTURE REFERENCE INTERNE: 12345 ... more text"
    txt_b = "Some text... Montant brut hors tva 123,45 ... more text"
    txt_unknown = "Hello world invoice"

    assert detect_type(txt_a) == "A"
    assert detect_type(txt_b) == "B"
    assert detect_type(txt_unknown) is None


# Test validation logic
def test_validate_invoice_type_a():
    # Valid Type A
    row_valid = {
        "N_FACTURE": "123456",
        "Date": "01/01/26",
        "N_CLIENT": "000123",
        "MONTANT_A_PAYER": "100.00",
        "TOTAUX_Mts_MARCH_HTVA": "80.00",
        "TOTAUX_Mts_TVA": "20.00",
    }
    status, notes = validate_invoice("TypeA", row_valid)
    assert status == "VALID"
    assert notes == ""

    # Missing required field
    row_missing = {
        "N_FACTURE": "123456",
        "Date": "01/01/26",
    }
    status, notes = validate_invoice("TypeA", row_missing)
    assert status == "SUSPICIOUS"
    assert "Missing required field" in notes

    # Invalid Date
    row_bad_date = {
        "N_FACTURE": "123456",
        "Date": "abc",
        "N_CLIENT": "000123",
        "MONTANT_A_PAYER": "100.00",
    }
    status, notes = validate_invoice("TypeA", row_bad_date)
    assert status == "SUSPICIOUS"
    assert "Invalid date format" in notes

    # Math mismatch
    row_bad_math = {
        "N_FACTURE": "123456",
        "Date": "01/01/26",
        "N_CLIENT": "000123",
        "MONTANT_A_PAYER": "120.00",
        "TOTAUX_Mts_MARCH_HTVA": "80.00",
        "TOTAUX_Mts_TVA": "20.00",
    }
    status, notes = validate_invoice("TypeA", row_bad_math)
    assert "Math check" in notes


def test_validate_invoice_type_b():
    # Valid Type B
    row_valid = {
        "FACTURE_N": "970584",
        "Date_facture": "08/06/2026",
        "Montant_a_payer": "362.42",
        "Montant_brut_hors_tva": "342.32",
        "Montant_TVA": "20.10",
    }
    status, notes = validate_invoice("TypeB", row_valid)
    assert status == "VALID"
    assert notes == ""

    # Missing required field
    row_missing = {
        "FACTURE_N": "970584",
    }
    status, notes = validate_invoice("TypeB", row_missing)
    assert status == "SUSPICIOUS"
    assert "Missing required field" in notes


# Test file discovery
def test_discover_files(tmp_path):
    # Setup mock files
    pdf1 = tmp_path / "test1.pdf"
    pdf2 = tmp_path / "test2.PDF"
    txt_file = tmp_path / "other.txt"
    pdf1.touch()
    pdf2.touch()
    txt_file.touch()

    # Discover in directory
    files = discover_files([str(tmp_path)])
    assert len(files) == 2
    assert pdf1 in files or pdf2 in files


# Test folder recursion
def test_discover_files_recursive(tmp_path):
    sub = tmp_path / "subdir"
    sub.mkdir()
    pdf1 = tmp_path / "test1.pdf"
    pdf2 = sub / "test2.pdf"
    pdf1.touch()
    pdf2.touch()

    # Recursive
    files_rec = discover_files([str(tmp_path)], recursive=True)
    assert len(files_rec) == 2

    # Non-recursive
    files_non = discover_files([str(tmp_path)], recursive=False)
    assert len(files_non) == 1


# Test extensibility / registering new parser
def test_extensibility():
    class TypeCParser(BaseParser):
        @property
        def sheet_name(self) -> str:
            return "TypeC"

        @property
        def columns(self) -> list:
            return ["Fichier", "SpecialField"]

        def detect(self, page1_text: str) -> bool:
            return "SPECIAL_MARKER_C" in page1_text

        def parse(self, pages_text: list, filename: str) -> dict:
            return {"Fichier": filename, "SpecialField": "FoundC"}

    # Register custom parser
    register_parser("TypeC", TypeCParser())

    # Verify registration
    assert "TypeC" in registry.get_all()

    # Test detection and processing
    pages = ["SPECIAL_MARKER_C\nContent"]
    kind, row, _ = registry.detect_and_process(pages, "mock.pdf")
    assert kind == "TypeC"
    assert row["SpecialField"] == "FoundC"


# Integration / Parser run tests using fixtures
def test_parser_integration():
    fixture_dir = Path(__file__).parent / "fixtures"
    type1_pdf = fixture_dir / "Type1.pdf"
    type2_pdf = fixture_dir / "Type2.pdf"

    assert type1_pdf.exists(), "Type1.pdf fixture is missing"
    assert type2_pdf.exists(), "Type2.pdf fixture is missing"

    # Process Type 1 (Type B layout)
    kind1, row1 = process(type1_pdf)
    assert kind1 == "TypeB"
    assert row1["FACTURE_N"] == "970584"
    assert row1["Date_facture"] == "08/06/2026"
    assert float(row1["Montant_a_payer"]) == 362.42

    # Process Type 2 (Type A layout)
    kind2, row2 = process(type2_pdf)
    assert kind2 == "TypeA"
    assert row2["N_FACTURE"] == "169901"
    assert row2["Date"] == "08/06/26"
    assert float(row2["MONTANT_A_PAYER"]) == 5845.33
