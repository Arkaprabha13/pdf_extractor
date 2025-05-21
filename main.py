"""
PDF Text and Table Extraction Tool
Version: 2.0
Date: 2025-05-21
Authors: Perplexity AI Research
"""

import os
import re
import json
import fitz  # PyMuPDF
import pytesseract
import pdfplumber
import camelot
import tabula
from pdf2image import convert_from_path
from typing import Dict, Any, List
from pypdf import PdfReader
from pypdf.errors import PdfReadError

class AdvancedPDFExtractor:
    """Advanced PDF extraction tool combining multiple parsing strategies"""
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.results = {
            "metadata": {},
            "headers": {},
            "content": [],
            "list_items": [],
            "tables": [],
            "figures": []
        }
        self.extraction_methods = []
        self.is_scanned = False
        self.ocr_engine = "Tesseract"
        
        # Configure Tesseract path if needed
        if os.name == 'nt':  # Windows
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    def extract(self) -> Dict[str, Any]:
        """Main extraction workflow"""
        self._extract_metadata()
        
        # Attempt extraction methods in priority order
        methods = [
            self._extract_with_pymupdf,
            self._extract_with_pdfplumber,
            self._extract_with_pypdf,
            self._ocr_fallback
        ]
        
        for method in methods:
            if method():
                break
                
        return self.results

    def _extract_metadata(self) -> None:
        """Extract PDF metadata using PyMuPDF"""
        with fitz.open(self.pdf_path) as doc:
            metadata = doc.metadata
            self.results["metadata"] = {
                "title": metadata.get("title", ""),
                "author": metadata.get("author", ""),
                "creator": metadata.get("creator", ""),
                "creation_date": metadata.get("creation_date", ""),
                "modification_date": metadata.get("mod_date", ""),
                "page_count": doc.page_count
            }

    def _extract_with_pymupdf(self) -> bool:
        """Text extraction using PyMuPDF (fitz)"""
        try:
            text_content = []
            with fitz.open(self.pdf_path) as doc:
                for page_num, page in enumerate(doc):
                    text = page.get_text("text", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_MEDIABOX_CLIP)
                    if text.strip():
                        text_content.append({
                            "page": page_num + 1,
                            "content": text,
                            "type": "text"
                        })
            
            if text_content:
                self.results["content"].extend(text_content)
                self.extraction_methods.append("PyMuPDF")
                self._extract_tables()
                return True
            return False
        except Exception as e:
            print(f"PyMuPDF extraction failed: {e}")
            return False

    def _extract_with_pdfplumber(self) -> bool:
        """Advanced extraction with PDFPlumber"""
        try:
            text_content = []
            with pdfplumber.open(self.pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text(layout=True, x_density=3, y_density=3)
                    if text.strip():
                        text_content.append({
                            "page": page_num + 1,
                            "content": text,
                            "type": "text"
                        })
                    
                    # Extract figures
                    for image in page.images:
                        self.results["figures"].append({
                            "page": page_num + 1,
                            "bbox": image["bbox"],
                            "width": image["width"],
                            "height": image["height"]
                        })
            
            if text_content:
                self.results["content"].extend(text_content)
                self.extraction_methods.append("PDFPlumber")
                self._extract_tables()
                return True
            return False
        except Exception as e:
            print(f"PDFPlumber extraction failed: {e}")
            return False

    def _extract_with_pypdf(self) -> bool:
        """Fallback extraction with pypdf"""
        try:
            text_content = []
            reader = PdfReader(self.pdf_path)
            for page_num, page in enumerate(reader.pages):
                text = page.extract_text(extraction_mode="layout", layout_mode_space_vertically=False)
                if text.strip():
                    text_content.append({
                        "page": page_num + 1,
                        "content": text,
                        "type": "text"
                    })
            
            if text_content:
                self.results["content"].extend(text_content)
                self.extraction_methods.append("PyPDF")
                self._extract_tables()
                return True
            return False
        except PdfReadError as e:
            print(f"Encrypted PDF detected: {e}")
            return False
        except Exception as e:
            print(f"PyPDF extraction failed: {e}")
            return False

    def _ocr_fallback(self) -> bool:
        """Final OCR fallback using Tesseract"""
        try:
            images = convert_from_path(self.pdf_path, dpi=300)
            text_content = []
            
            for page_num, img in enumerate(images):
                text = pytesseract.image_to_string(img, config='--psm 6')
                if text.strip():
                    text_content.append({
                        "page": page_num + 1,
                        "content": text,
                        "type": "ocr_text"
                    })
            
            if text_content:
                self.results["content"].extend(text_content)
                self.extraction_methods.append("Tesseract OCR")
                self.is_scanned = True
                return True
            return False
        except Exception as e:
            print(f"OCR extraction failed: {e}")
            return False

    def _extract_tables(self) -> None:
        """Multi-method table extraction"""
        # Try Camelot first
        try:
            camelot_tables = camelot.read_pdf(self.pdf_path, flavor='lattice', pages='all')
            for table in camelot_tables:
                if table.df.empty:
                    continue
                self.results["tables"].append({
                    "method": "Camelot",
                    "page": table.page,
                    "data": table.df.replace(r'^\s*$', None, regex=True).to_dict(orient='records')
                })
        except Exception as e:
            print(f"Camelot table extraction failed: {e}")

        # Try Tabula as fallback
        try:
            tabula_tables = tabula.read_pdf(self.pdf_path, pages='all', multiple_tables=True)
            for i, table in enumerate(tabula_tables):
                if not table.empty:
                    self.results["tables"].append({
                        "method": "Tabula",
                        "page": i+1,
                        "data": table.fillna('').to_dict(orient='records')
                    })
        except Exception as e:
            print(f"Tabula table extraction failed: {e}")

    def _process_content(self) -> None:
        """Post-process extracted content"""
        # Header detection patterns
        header_patterns = [
            re.compile(r'^\s*(?:[A-Z][A-Z\s]+:?\s*)$'),
            re.compile(r'^\s*\d+\.\s+[A-Z][a-zA-Z\s]+'),
            re.compile(r'^\s*§\s*\d+\.\s.*'),
            re.compile(r'^\s*[IVX]+\.\s.*')
        ]

        current_header = "Main Content"
        for item in self.results["content"]:
            lines = item["content"].split('\n')
            for line in lines:
                line = line.strip()
                if any(pattern.match(line) for pattern in header_patterns):
                    current_header = line
                    self.results["headers"][current_header] = []
                else:
                    if current_header in self.results["headers"]:
                        self.results["headers"][current_header].append(line)
                    else:
                        self.results["headers"]["Main Content"] = [line]

def main():
    """Command line interface"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Advanced PDF Text Extraction Tool',
        epilog='Supports text extraction from both digital and scanned PDFs'
    )
    parser.add_argument('pdf_path', help='Path to input PDF file')
    parser.add_argument('-o', '--output', help='Output JSON file path')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose output')
    
    args = parser.parse_args()
    
    try:
        extractor = AdvancedPDFExtractor(args.pdf_path)
        results = extractor.extract()
        extractor._process_content()
        
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            print(f"Results saved to {args.output}")
        else:
            print(json.dumps(results, indent=2, ensure_ascii=False))
            
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        exit(1)

if __name__ == "__main__":
    main()
