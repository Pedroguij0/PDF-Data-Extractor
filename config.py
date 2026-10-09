from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
INPUT_PDF_DIR = BASE_DIR
OUTPUT_DIR = BASE_DIR / "output"
SIGNATURES_DIR = OUTPUT_DIR / "signatures"
OUTPUT_EXCEL_FILE = OUTPUT_DIR / "Extracted_Form.xlsx"

#"hyperlink": Save the signature png in /signature and insert an clickable hyperlink in excel cell.
#"base64": Save a complete string in excel cell and not insert any file in /signature.
#"both": Save both link and string in excel cell and /signature.

SIGNATURE_CAPTURE_MODE = "both"
#Dictionary that contain all form fileds and transform they into a excel column title
FIELD_MAPPING = {}
SYSTEM_COLUMNS = {
    "source_file":"Source PDF File",
    "signature_status":"Signature Detected",
    "signature_link":"Signature Artifact Link",
    "signature_base64":"Signature Base64 String"
}   