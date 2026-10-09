import sys
import time
from pathlib import Path
from typing import List
import config
from excel_builder import ExcelBuilder
from extractor import PDFFormExtractor

def discover_pdf_files(target_dir: Path) -> List[Path]:
    return sorted([file for file in target_dir.glob("*.pdf") if file.is_file()])

def run_pipeline():
    start_time = time.time()
    print("\n"+"="*65)
    print("  AUTOMATED FILLABLE PDF TO EXCEL EXTRACTION PIPELINE")
    print("=" * 65)
    print(f"Search Directory: {config.INPUT_PDF_DIR.resolve()}")
    pdf_files = discover_pdf_files(config.INPUT_PDF_DIR)
    total_files = len(pdf_files)
    if total_files == 0:
        print(f"[WARNING] No pdf file found in {config.INPUT_PDF_DIR}.")
        sys.exit(1)
    
    print(f"Total PDF Forms: {total_files}")
    print("="*65)
    records = []
    for idx, pdf_path in enumerate(pdf_files, 1):
        print(f"[{idx:03d}/{total_files:03d}] Processing: {pdf_path.name}...", end="")
        try:
            extractor = PDFFormExtractor(pdf_path)
            data = extractor.extract_all()
            records.append(data)
            print("FINISHED")
        except Exception as e:
            print(f"Failed to extract data from pdf {pdf_path.name}: {e}")
        
    print("="*65)
    print("Generating Stylized Excel Worksheet and Linking Signatures...")
    builder = ExcelBuilder(records, output_path=config.OUTPUT_EXCEL_FILE)
    output_excel = builder.build()
    elapsed_time = time.time() - start_time
    print(f"SUCCESS! Worksheet successfully generated with {len(records)} records:")
    print("-"*10+f"Signatures saved in: {config.SIGNATURES_DIR}")
    print(f"Total Generating Time: {elapsed_time:.2f} seconds")
    print("="*65 +"\n")

if __name__ == "__main__":
    run_pipeline()