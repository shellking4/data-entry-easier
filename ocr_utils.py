import pytesseract
from pdf2image import convert_from_path
import os
import tempfile
from PIL import Image

def needs_ocr(pdf_path):
    """
    Check if a PDF needs OCR by attempting to extract text from the first page.
    This is a heuristic: if we get very little text, we assume it's scanned.
    """
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return False
            
            # Check first page
            text = pdf.pages[0].extract_text()
            if not text or len(text.strip()) < 10:
                return True
            
            # Also check if there are any tables detected
            tables = pdf.pages[0].extract_tables()
            if not tables:
                # If no text AND no tables, definitely needs OCR
                return True
                
    except Exception as e:
        print(f"Error checking if OCR is needed: {e}")
        # If we can't read it, maybe it's corrupted or needs OCR? 
        # Let's assume safely that if we can't read it normally, we might try OCR or just fail later.
        pass
        
    return False

def convert_to_searchable_pdf(pdf_path, output_path=None):
    """
    Convert a scanned PDF to a searchable PDF using Tesseract.
    Returns the path to the new PDF.
    """
    try:
        print(f"Converting {pdf_path} to images...")
        # Convert PDF to images
        images = convert_from_path(pdf_path)
        
        if not images:
            raise ValueError("No images extracted from PDF")
            
        print(f"Running OCR on {len(images)} pages...")
        
        # Create a new PDF with OCR data
        pdf_writer = pytesseract.image_to_pdf_or_hocr(images[0], extension='pdf')
        
        # If there are multiple pages, we need to handle them. 
        # pytesseract.image_to_pdf_or_hocr typically handles one image.
        # We might need to append them.
        # Actually, it's better to use a library that can combine them, or just loop.
        # But pytesseract can take a list of images? No, usually one.
        
        # Better approach: Use pytesseract to get PDF bytes for each page and merge them.
        # Or use a simpler approach: save images to a list and use tesseract to process them?
        # Tesseract CLI can handle multiple images.
        
        # Let's use a simpler approach for now:
        # We will use `pdf2image` to get images, then `pytesseract` to get text? 
        # No, we need a searchable PDF (image + text layer).
        
        # We can use `pytesseract.image_to_pdf_or_hocr` for each page and merge.
        # But merging PDFs requires another lib (like PyPDF2 or pypdf).
        # We have `pdfplumber` which reads, but doesn't write.
        # We have `pdf2image`.
        
        # Let's try to use `pytesseract` to process the whole PDF if possible? No.
        
        # Alternative: iterate pages, get PDF bytes, merge.
        # Since we don't have PyPDF2 in requirements (only pypdfium2, pdfplumber), 
        # we might need to add `pypdf` or `PyPDF2` to merge.
        # Wait, `pdfplumber` uses `pdfminer.six`.
        
        # Let's check what we have. `pypdfium2` is there.
        # Maybe we can just return the text? 
        # But `app.py` expects a PDF file path to pass to `extract_pdf_data`.
        # So we MUST generate a PDF file.
        
        # Let's add `pypdf` to requirements or use `pdfplumber`'s underlying tools? No.
        # Let's just use `pytesseract` to get text and return text?
        # No, `extract_pdf_data` in `app.py` uses `pdfplumber` to extract TABLES.
        # Tesseract can output a PDF with text layer.
        
        # If we process page by page, we get multiple PDF bytes.
        # We need to merge them.
        
        # Let's use `pypdf` (it's standard and lightweight). 
        # I'll add it to requirements.
        
        from pypdf import PdfWriter, PdfReader
        import io
        
        merger = PdfWriter()
        
        for i, image in enumerate(images):
            print(f"Processing page {i+1}...")
            pdf_bytes = pytesseract.image_to_pdf_or_hocr(image, extension='pdf')
            merger.append(PdfReader(io.BytesIO(pdf_bytes)))
            
        if output_path is None:
            fd, output_path = tempfile.mkstemp(suffix='.pdf')
            os.close(fd)
            
        merger.write(output_path)
        merger.close()
        
        print(f"Searchable PDF saved to {output_path}")
        return output_path

    except Exception as e:
        print(f"OCR failed: {e}")
        raise e
