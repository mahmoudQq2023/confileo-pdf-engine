#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to Word Converter with Full Arabic RTL Support
===================================================
Uses free open-source libraries:
- PyMuPDF (fitz): PDF text and image extraction
- python-docx: Generate Word documents (Arabic kept in LOGICAL order; Word
  shapes + lays it out RTL itself — we never reshape or bidi-reorder)

Author: Confileo Team
License: MIT
"""

import fitz  # PyMuPDF
import unicodedata
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from PIL import Image
import re
import io
import sys
import json
import os
import argparse
import logging
from typing import Optional, List, Tuple, Dict, Any
from dataclasses import dataclass
from pathlib import Path

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

# Default Arabic fonts
ARABIC_FONTS = [
    'Arial',
    'Tahoma', 
    'Traditional Arabic',
    'Simplified Arabic',
    'Sakkal Majalla',
    'Amiri',
    'Noto Naskh Arabic',
]


@dataclass
class ConversionOptions:
    """Options for PDF to Word conversion"""
    preserve_images: bool = True
    preserve_fonts: bool = True
    detect_rtl: bool = True
    default_font: str = 'Arial'
    default_font_size: float = 12.0
    image_quality: int = 95
    max_image_width: float = 6.0  # inches
    page_margin: float = 1.0  # inches


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


def process_arabic_text(text: str, force_rtl: bool = False) -> Tuple[str, bool]:
    """Return (logical_text, is_rtl) for writing into a .docx.

    We do NOT reshape or bidi-reorder: Word shapes logical Unicode and lays it out
    RTL itself, so reshape() + get_display() would bake visual-order presentation
    forms into the document and render as broken, reversed Arabic. We only
    NFKC-normalise any pre-shaped presentation forms back to base letters.
    """
    if not text:
        return text, False

    if not is_arabic_text(text) and not force_rtl:
        return text, False

    return unicodedata.normalize("NFKC", text), True


def set_paragraph_rtl(paragraph) -> None:
    """Set paragraph direction to RTL"""
    pPr = paragraph._p.get_or_add_pPr()
    
    # Set bidi (bidirectional)
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    
    # Set text alignment to right
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def set_run_rtl(run) -> None:
    """Set run (text span) direction to RTL"""
    rPr = run._r.get_or_add_rPr()
    
    # Set RTL
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)


def set_document_rtl(doc: Document) -> None:
    """Set document default direction to RTL"""
    for section in doc.sections:
        sectPr = section._sectPr
        
        # Set bidi for section
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        sectPr.append(bidi)
        
        # Set RTL gutter
        rtlGutter = OxmlElement('w:rtlGutter')
        rtlGutter.set(qn('w:val'), '1')
        sectPr.append(rtlGutter)


def set_table_rtl(table) -> None:
    """Set table direction to RTL"""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
    
    # Set bidi for table
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    tblPr.append(bidi)


def extract_font_info(span: Dict[str, Any]) -> Tuple[str, float, Optional[Tuple[int, int, int]]]:
    """Extract font information from a text span"""
    font_name = span.get('font', 'Arial')
    font_size = span.get('size', 12.0)
    
    # Extract color if available
    color = None
    if 'color' in span:
        c = span['color']
        if isinstance(c, int):
            # Convert integer color to RGB
            color = ((c >> 16) & 255, (c >> 8) & 255, c & 255)
    
    return font_name, font_size, color


def extract_images_from_page(page, pdf_doc) -> List[Tuple[bytes, str, Dict]]:
    """Extract all images from a PDF page"""
    images = []
    
    try:
        image_list = page.get_images(full=True)
        
        for img_index, img_info in enumerate(image_list):
            try:
                xref = img_info[0]
                base_image = pdf_doc.extract_image(xref)
                
                if base_image:
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    # Get image rectangle if available
                    img_rect = None
                    for img in page.get_image_info():
                        if img.get('xref') == xref:
                            img_rect = img.get('bbox')
                            break
                    
                    images.append((image_bytes, image_ext, {
                        'xref': xref,
                        'rect': img_rect,
                        'width': base_image.get('width', 0),
                        'height': base_image.get('height', 0)
                    }))
            except Exception as e:
                logger.warning(f"Failed to extract image {img_index}: {e}")
                
    except Exception as e:
        logger.warning(f"Failed to get images from page: {e}")
    
    return images


def add_image_to_doc(doc: Document, image_bytes: bytes, image_ext: str, 
                     max_width: float = 6.0) -> bool:
    """Add an image to the Word document"""
    try:
        # Load image to get dimensions
        img = Image.open(io.BytesIO(image_bytes))
        width, height = img.size
        
        # Calculate aspect ratio
        aspect_ratio = height / width if width > 0 else 1
        
        # Calculate display width (max 6 inches)
        display_width = min(max_width, width / 96)  # assuming 96 DPI
        
        # Add paragraph for image
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run = para.add_run()
        run.add_picture(io.BytesIO(image_bytes), width=Inches(display_width))
        
        return True
    except Exception as e:
        logger.warning(f"Failed to add image: {e}")
        return False


def process_text_block(block: Dict[str, Any], doc: Document, 
                       options: ConversionOptions) -> Tuple[int, bool]:
    """
    Process a text block and add it to the document
    
    Returns:
        Tuple of (paragraphs_added, contains_arabic)
    """
    paragraphs_added = 0
    contains_arabic = False
    
    if block.get('type') != 0:  # Not a text block
        return paragraphs_added, contains_arabic
    
    lines = block.get('lines', [])
    
    for line in lines:
        spans = line.get('spans', [])
        if not spans:
            continue
        
        # Combine all spans in the line
        line_text = ''
        line_has_arabic = False
        font_info = None
        
        for span in spans:
            span_text = span.get('text', '')
            line_text += span_text
            
            if is_arabic_text(span_text):
                line_has_arabic = True
                contains_arabic = True
            
            # Get font info from first span
            if font_info is None:
                font_info = extract_font_info(span)
        
        if not line_text.strip():
            continue
        
        # Process Arabic text if detected
        if line_has_arabic and options.detect_rtl:
            processed_text, is_rtl = process_arabic_text(line_text)
        else:
            processed_text = line_text
            is_rtl = False
        
        # Create paragraph
        para = doc.add_paragraph()
        
        if is_rtl:
            set_paragraph_rtl(para)
        
        # Add text run
        run = para.add_run(processed_text)
        
        if is_rtl:
            set_run_rtl(run)
        
        # Apply font formatting
        if font_info and options.preserve_fonts:
            font_name, font_size, color = font_info
            
            run.font.size = Pt(font_size)
            
            # Use Arabic-compatible font
            if is_rtl:
                run.font.name = options.default_font
                run._element.rPr.rFonts.set(qn('w:cs'), options.default_font)
            else:
                run.font.name = font_name
            
            # Apply color if available
            if color:
                run.font.color.rgb = RGBColor(*color)
        
        paragraphs_added += 1
    
    return paragraphs_added, contains_arabic


def convert_pdf_to_word(pdf_path: str, output_path: str, 
                        options: Optional[ConversionOptions] = None) -> ConversionResult:
    """
    Convert PDF to Word document with Arabic RTL support
    
    Args:
        pdf_path: Path to input PDF file
        output_path: Path for output DOCX file
        options: Conversion options
        
    Returns:
        ConversionResult with details about the conversion
    """
    import time
    start_time = time.time()
    
    if options is None:
        options = ConversionOptions()
    
    result = ConversionResult(success=False)
    
    try:
        # Validate input file
        if not os.path.exists(pdf_path):
            result.error = f"PDF file not found: {pdf_path}"
            return result
        
        # Open PDF document
        pdf_doc = fitz.open(pdf_path)
        result.page_count = len(pdf_doc)
        
        logger.info(f"Processing PDF: {pdf_path} ({result.page_count} pages)")
        
        # Create Word document
        doc = Document()
        
        # Set page margins
        for section in doc.sections:
            section.left_margin = Inches(options.page_margin)
            section.right_margin = Inches(options.page_margin)
            section.top_margin = Inches(options.page_margin)
            section.bottom_margin = Inches(options.page_margin)
        
        total_text_blocks = 0
        total_images = 0
        has_arabic = False
        
        # Process each page
        for page_num, page in enumerate(pdf_doc):
            logger.info(f"Processing page {page_num + 1}/{result.page_count}")
            
            # Extract text blocks
            blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
            
            for block in blocks:
                if block.get('type') == 0:  # Text block
                    paragraphs, block_has_arabic = process_text_block(block, doc, options)
                    total_text_blocks += 1
                    if block_has_arabic:
                        has_arabic = True
                        
                elif block.get('type') == 1 and options.preserve_images:  # Image block
                    # Extract and add image
                    pass  # Images handled separately
            
            # Extract and add images
            if options.preserve_images:
                images = extract_images_from_page(page, pdf_doc)
                for img_bytes, img_ext, img_info in images:
                    if add_image_to_doc(doc, img_bytes, img_ext, options.max_image_width):
                        total_images += 1
            
            # Add page break (except for last page)
            if page_num < result.page_count - 1:
                doc.add_page_break()
        
        # Set document RTL if Arabic was detected
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
        logger.info(f"  - Text blocks: {total_text_blocks}")
        logger.info(f"  - Images: {total_images}")
        logger.info(f"  - Arabic detected: {has_arabic}")
        logger.info(f"  - Time: {result.processing_time:.2f}s")
        
    except Exception as e:
        result.error = str(e)
        logger.error(f"Conversion failed: {e}")
    
    return result


def main():
    """Main entry point for CLI usage"""
    parser = argparse.ArgumentParser(
        description='Convert PDF to Word with Arabic RTL support'
    )
    parser.add_argument('input', help='Input PDF file path')
    parser.add_argument('output', help='Output DOCX file path')
    parser.add_argument('--no-images', action='store_true', 
                        help='Skip image extraction')
    parser.add_argument('--no-rtl', action='store_true',
                        help='Disable RTL detection')
    parser.add_argument('--font', default='Arial',
                        help='Default font for Arabic text')
    parser.add_argument('--font-size', type=float, default=12.0,
                        help='Default font size')
    parser.add_argument('--json', action='store_true',
                        help='Output result as JSON')
    
    args = parser.parse_args()
    
    options = ConversionOptions(
        preserve_images=not args.no_images,
        detect_rtl=not args.no_rtl,
        default_font=args.font,
        default_font_size=args.font_size
    )
    
    result = convert_pdf_to_word(args.input, args.output, options)
    
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
            print(f"   Arabic RTL: {'Yes' if result.arabic_detected else 'No'}")
            print(f"   Time: {result.processing_time:.2f}s")
        else:
            print(f"❌ Conversion failed: {result.error}")
            sys.exit(1)


if __name__ == "__main__":
    main()
