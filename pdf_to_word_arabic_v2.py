#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to Word Converter with Arabic OCR Support
==============================================
Uses Tesseract OCR for accurate Arabic text extraction
This approach bypasses PDF presentation forms issues

Dependencies:
- pdf2image: Convert PDF pages to images
- pytesseract: OCR engine wrapper
- python-docx: Generate Word documents
- Pillow: Image processing

Author: Confileo Team
License: MIT
"""

import fitz  # PyMuPDF for images
import pytesseract
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io
import sys
import json
import os
import argparse
import logging
from typing import Optional, List, Tuple, Dict, Any
from dataclasses import dataclass

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Arabic Unicode ranges
ARABIC_RANGES = [
    (0x0600, 0x06FF),  # Arabic
    (0x0750, 0x077F),  # Arabic Supplement
    (0x08A0, 0x08FF),  # Arabic Extended-A
    (0xFB50, 0xFDFF),  # Arabic Presentation Forms-A
    (0xFE70, 0xFEFF),  # Arabic Presentation Forms-B
]


@dataclass
class ConversionResult:
    """Result of PDF to Word conversion"""
    success: bool
    output_path: Optional[str] = None
    page_count: int = 0
    text_blocks: int = 0
    images_extracted: int = 0
    arabic_detected: bool = False
    error: Optional[str] = None
    processing_time: float = 0.0


def is_arabic_char(char: str) -> bool:
    """Check if a character is Arabic"""
    code = ord(char)
    for start, end in ARABIC_RANGES:
        if start <= code <= end:
            return True
    return False


def is_arabic_text(text: str) -> bool:
    """Check if text contains Arabic characters"""
    if not text:
        return False
    return any(is_arabic_char(c) for c in text)


def get_arabic_ratio(text: str) -> float:
    """Get the ratio of Arabic characters in text"""
    if not text:
        return 0.0
    arabic_count = sum(1 for c in text if is_arabic_char(c))
    alpha_count = sum(1 for c in text if c.isalpha())
    return arabic_count / alpha_count if alpha_count > 0 else 0.0


def set_paragraph_rtl(paragraph) -> None:
    """Set paragraph direction to RTL"""
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def set_run_rtl(run) -> None:
    """Set run direction to RTL"""
    rPr = run._r.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)


def set_document_rtl(doc: Document) -> None:
    """Set document default direction to RTL"""
    for section in doc.sections:
        sectPr = section._sectPr
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        sectPr.append(bidi)


def extract_text_with_ocr(page_image: Image.Image, lang: str = 'ara+eng') -> str:
    """
    Extract text from image using Tesseract OCR
    """
    try:
        # Configure Tesseract for better Arabic support
        custom_config = r'--oem 3 --psm 6'
        
        # Run OCR
        text = pytesseract.image_to_string(
            page_image, 
            lang=lang,
            config=custom_config
        )
        
        return text.strip()
    except Exception as e:
        logger.warning(f"OCR failed: {e}")
        return ""


def pdf_page_to_image(page, dpi: int = 200) -> Image.Image:
    """Convert PDF page to PIL Image using PyMuPDF"""
    # Create high-resolution image
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    
    # Convert to PIL Image
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))
    
    return img


def convert_pdf_to_word_ocr(pdf_path: str, output_path: str, 
                            use_ocr: bool = True,
                            preserve_images: bool = True,
                            ocr_lang: str = 'ara+eng') -> ConversionResult:
    """
    Convert PDF to Word using OCR for Arabic text
    """
    import time
    start_time = time.time()
    
    result = ConversionResult(success=False)
    
    try:
        # Validate input
        if not os.path.exists(pdf_path):
            result.error = f"PDF file not found: {pdf_path}"
            return result
        
        # Open PDF
        pdf_doc = fitz.open(pdf_path)
        result.page_count = len(pdf_doc)
        
        logger.info(f"Processing PDF: {pdf_path} ({result.page_count} pages)")
        
        # Create Word document
        doc = Document()
        
        # Set margins
        for section in doc.sections:
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
        
        has_arabic = False
        total_text_blocks = 0
        total_images = 0
        
        for page_num, page in enumerate(pdf_doc):
            logger.info(f"Processing page {page_num + 1}/{result.page_count}")
            
            if use_ocr:
                # Convert page to image
                page_image = pdf_page_to_image(page, dpi=200)
                
                # Extract text using OCR
                text = extract_text_with_ocr(page_image, lang=ocr_lang)
                
                if text:
                    # Check for Arabic
                    is_rtl = is_arabic_text(text)
                    if is_rtl:
                        has_arabic = True
                    
                    # Reconstruct paragraphs intelligently
                    lines = text.split('\n')
                    paragraphs = []
                    current_para = ""
                    
                    for line in lines:
                        line = line.strip()
                        if not line:
                            if current_para:
                                paragraphs.append(current_para)
                                current_para = ""
                            continue
                        
                        if current_para:
                            # If previous line doesn't end with punctuation, merge with space
                            if current_para[-1] not in '.!?،؛:':
                                current_para += " " + line
                            else:
                                paragraphs.append(current_para)
                                current_para = line
                        else:
                            current_para = line
                            
                    if current_para:
                        paragraphs.append(current_para)

                    for para_text in paragraphs:
                        para = doc.add_paragraph()
                        line_is_arabic = is_arabic_text(para_text)
                        
                        if line_is_arabic:
                            set_paragraph_rtl(para)
                        
                        run = para.add_run(para_text)
                        run.font.size = Pt(12)
                        run.font.name = 'Arial'
                        run._element.rPr.rFonts.set(qn('w:cs'), 'Arial')
                        
                        if line_is_arabic:
                            set_run_rtl(run)
                        
                        total_text_blocks += 1
            else:
                # Fallback: Use PyMuPDF text extraction
                blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
                
                for block in blocks:
                    if block.get('type') == 0:  # Text block
                        for line in block.get('lines', []):
                            line_text = ''.join(span.get('text', '') for span in line.get('spans', []))
                            line_text = line_text.strip()
                            
                            if not line_text:
                                continue
                            
                            para = doc.add_paragraph()
                            is_rtl = is_arabic_text(line_text)
                            
                            if is_rtl:
                                has_arabic = True
                                set_paragraph_rtl(para)
                            
                            run = para.add_run(line_text)
                            run.font.size = Pt(12)
                            run.font.name = 'Arial'
                            
                            if is_rtl:
                                set_run_rtl(run)
                            
                            total_text_blocks += 1
            
            # Extract images if requested
            if preserve_images:
                try:
                    image_list = page.get_images(full=True)
                    for img_index, img_info in enumerate(image_list):
                        try:
                            xref = img_info[0]
                            base_image = pdf_doc.extract_image(xref)
                            if base_image:
                                image_bytes = base_image["image"]
                                
                                para = doc.add_paragraph()
                                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                run = para.add_run()
                                
                                # Calculate width (max 5 inches)
                                img = Image.open(io.BytesIO(image_bytes))
                                width, height = img.size
                                display_width = min(5, width / 96)
                                
                                run.add_picture(io.BytesIO(image_bytes), width=Inches(display_width))
                                total_images += 1
                        except Exception as e:
                            logger.warning(f"Failed to extract image: {e}")
                except Exception as e:
                    logger.warning(f"Failed to get images: {e}")
            
            # Add page break (except last page)
            if page_num < result.page_count - 1:
                doc.add_page_break()
        
        # Set document RTL if Arabic detected
        if has_arabic:
            set_document_rtl(doc)
            logger.info("Arabic text detected - document set to RTL")
        
        # Save document
        doc.save(output_path)
        pdf_doc.close()
        
        # Update result
        result.success = True
        result.output_path = output_path
        result.text_blocks = total_text_blocks
        result.images_extracted = total_images
        result.arabic_detected = has_arabic
        result.processing_time = time.time() - start_time
        
        logger.info(f"Conversion complete: {output_path}")
        logger.info(f"  - Pages: {result.page_count}")
        logger.info(f"  - Text blocks: {total_text_blocks}")
        logger.info(f"  - Images: {total_images}")
        logger.info(f"  - Arabic detected: {has_arabic}")
        logger.info(f"  - Time: {result.processing_time:.2f}s")
        
    except Exception as e:
        result.error = str(e)
        logger.error(f"Conversion failed: {e}")
    
    return result


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Convert PDF to Word with Arabic OCR support'
    )
    parser.add_argument('input', help='Input PDF file path')
    parser.add_argument('output', help='Output DOCX file path')
    parser.add_argument('--no-ocr', action='store_true',
                        help='Disable OCR (use direct text extraction)')
    parser.add_argument('--no-images', action='store_true',
                        help='Skip image extraction')
    parser.add_argument('--lang', default='ara+eng',
                        help='OCR language (default: ara+eng)')
    parser.add_argument('--json', action='store_true',
                        help='Output result as JSON')
    
    args = parser.parse_args()
    
    result = convert_pdf_to_word_ocr(
        args.input,
        args.output,
        use_ocr=not args.no_ocr,
        preserve_images=not args.no_images,
        ocr_lang=args.lang
    )
    
    if args.json:
        output = {
            'success': result.success,
            'output_path': result.output_path,
            'page_count': result.page_count,
            'text_blocks': result.text_blocks,
            'images_extracted': result.images_extracted,
            'arabic_detected': result.arabic_detected,
            'processing_time': result.processing_time,
            'error': result.error
        }
        print(json.dumps(output, ensure_ascii=False))
    else:
        if result.success:
            print(f"✅ Conversion successful: {result.output_path}")
            print(f"   Pages: {result.page_count}")
            print(f"   Text blocks: {result.text_blocks}")
            print(f"   Images: {result.images_extracted}")
            print(f"   Arabic: {'Yes' if result.arabic_detected else 'No'}")
            print(f"   Time: {result.processing_time:.2f}s")
        else:
            print(f"❌ Conversion failed: {result.error}")
            sys.exit(1)


if __name__ == "__main__":
    main()
