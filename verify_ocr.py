import ocr_utils
import os

pdf_path = "fa.pdf"

print(f"Checking if {pdf_path} needs OCR...")
needs_ocr = ocr_utils.needs_ocr(pdf_path)
print(f"Needs OCR: {needs_ocr}")

# Force conversion to test the pipeline
print("Forcing conversion to searchable PDF...")
try:
    output = ocr_utils.convert_to_searchable_pdf(pdf_path)
    print(f"Conversion successful. Output: {output}")
    print(f"Output exists: {os.path.exists(output)}")
    
    # Clean up
    if os.path.exists(output):
        os.unlink(output)
        print("Cleaned up output file.")
except Exception as e:
    print(f"Conversion failed: {e}")
