"""
RTL Converter Module
Converts PowerPoint slide layout from Left-to-Right (LTR) to Right-to-Left (RTL)
This is the most complex module as it requires XML manipulation
"""
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from lxml import etree
from typing import Any
import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import Config
from utils.logger import setup_logger

logger = setup_logger(__name__)

# XML namespaces for PowerPoint
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

def flip_to_rtl_layout(input_path: str, output_path: str) -> None:
    """
    Convert PowerPoint slide from LTR to RTL layout

    This function prepares the slide for RTL content by:
    - Setting paragraph-level RTL direction (for proper text rendering)
    - Text alignment (RIGHT)
    - Shape positions (mirrored)
    - Reading order

    NOTE: We do NOT set slide-level RTL as it can cause character reversal
    with Arabic text that's already in correct logical order from AI translation.

    Args:
        input_path: Path to input .pptx file
        output_path: Path to save RTL-converted .pptx file
    """
    logger.info(f"Converting slide to RTL layout: {input_path}")

    try:
        # Load presentation
        prs = Presentation(input_path)

        # Process all slides (typically just one for MVP)
        for slide_idx, slide in enumerate(prs.slides):
            logger.info(f"Processing slide {slide_idx + 1}/{len(prs.slides)}")

            # Step 1: Flip shape positions horizontally (mirror the slide)
            slide_width = prs.slide_width
            for shape in slide.shapes:
                _flip_shape_position(shape, slide_width)

            # Step 2: Set text alignment to RIGHT and paragraph-level RTL
            for shape in slide.shapes:
                if shape.has_text_frame:
                    _set_text_rtl_and_alignment(shape)

            # Step 3: DO NOT set Arabic font here - let text_replacer preserve original formatting

        # Save modified presentation
        prs.save(output_path)
        logger.info(f"RTL conversion complete. Saved to: {output_path}")

    except Exception as e:
        logger.error(f"Error during RTL conversion: {str(e)}", exc_info=True)
        raise

def _set_slide_rtl_property(slide) -> None:
    """
    Set PowerPoint's built-in RTL property at the SLIDE level

    This is the correct way to enable RTL layout as per PowerPoint's design.
    Setting rtl="1" on the slide's p:cSld element activates PowerPoint's
    native right-to-left layout mode, which automatically:
    - Flips the reading order
    - Mirrors the slide layout
    - Sets text direction to RTL

    Args:
        slide: PowerPoint slide object
    """
    try:
        # Access slide XML element
        slide_element = slide._element

        # Find p:cSld (common slide data) element
        cSld = slide_element.find('.//p:cSld', namespaces=NAMESPACES)

        if cSld is not None:
            # Set RTL attribute on the slide (PowerPoint's built-in property)
            cSld.set('rtl', '1')  # 1 = RTL, 0 = LTR
            logger.info("✓ Set PowerPoint's built-in RTL property at slide level")
        else:
            logger.warning("Could not find p:cSld element to set RTL property")

    except Exception as e:
        logger.error(f"Error setting slide RTL property: {str(e)}")
        raise

def _set_text_rtl_and_alignment(shape) -> None:
    """
    Set text alignment to RIGHT and REMOVE any RTL attributes

    This sets:
    - RIGHT alignment (visual alignment to right side)
    - REMOVES rtl="1" attributes (they cause text reversal)

    NOTE: We REMOVE rtl="1" because Arabic text from LLM translation
    is already in correct logical order. The rtl="1" attribute would cause the text
    to be reversed again when displayed, making it unreadable.

    Args:
        shape: PowerPoint shape with text_frame
    """
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame

    for paragraph in text_frame.paragraphs:
        try:
            # Set RIGHT alignment
            paragraph.alignment = PP_ALIGN.RIGHT

            # REMOVE any existing RTL attribute
            _remove_rtl_from_paragraph(paragraph)
        except Exception as e:
            logger.warning(f"Could not set alignment for paragraph: {str(e)}")

def _set_text_alignment_right(shape) -> None:
    """
    Set text alignment to RIGHT for all paragraphs in a shape

    This complements the slide-level RTL property by ensuring
    text is visually aligned to the right side.

    Args:
        shape: PowerPoint shape with text_frame
    """
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame

    for paragraph in text_frame.paragraphs:
        try:
            paragraph.alignment = PP_ALIGN.RIGHT
        except Exception as e:
            logger.warning(f"Could not set alignment for paragraph: {str(e)}")

def _set_rtl_text_direction(shape) -> None:
    """
    Set text direction to RTL and align text to the right

    This function modifies:
    - Paragraph alignment (RIGHT)
    - Paragraph RTL property via XML (rtl="1")

    Args:
        shape: PowerPoint shape with text_frame
    """
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame

    for paragraph in text_frame.paragraphs:
        # Set alignment to RIGHT for RTL appearance
        paragraph.alignment = PP_ALIGN.RIGHT

        # Set RTL direction via XML manipulation
        try:
            _set_rtl_via_xml(paragraph)
        except Exception as e:
            logger.warning(f"Could not set RTL via XML: {str(e)}")

def _remove_rtl_from_paragraph(paragraph) -> None:
    """
    REMOVE RTL property from paragraph XML

    This removes the rtl="1" attribute that causes text reversal.
    Arabic text from LLM is already in correct logical order, so rtl="1"
    would reverse it and make it unreadable.

    Args:
        paragraph: PowerPoint paragraph object
    """
    try:
        # Get paragraph XML element
        p_element = paragraph._element

        # Get <a:pPr> (paragraph properties)
        pPr = p_element.find('.//a:pPr', namespaces=NAMESPACES)

        if pPr is not None and 'rtl' in pPr.attrib:
            # Remove RTL attribute
            del pPr.attrib['rtl']
            logger.debug("Removed RTL attribute from paragraph")

    except Exception as e:
        logger.warning(f"Could not remove RTL attribute: {str(e)}")

def _set_rtl_via_xml(paragraph) -> None:
    """
    Set RTL property in paragraph XML

    PowerPoint stores RTL information in the paragraph properties (pPr) XML element.
    We need to add rtl="1" attribute to the <a:pPr> element.

    Args:
        paragraph: PowerPoint paragraph object
    """
    # Get paragraph XML element
    p_element = paragraph._element

    # Get or create <a:pPr> (paragraph properties)
    pPr = p_element.find('.//a:pPr', namespaces=NAMESPACES)

    if pPr is None:
        # Create <a:pPr> if it doesn't exist
        # Insert before <a:r> (run) elements
        pPr = etree.Element(f"{{{NAMESPACES['a']}}}pPr")
        # Insert as first child
        p_element.insert(0, pPr)

    # Set RTL attribute
    pPr.set('rtl', '1')  # 1 = RTL, 0 = LTR

    logger.debug(f"Set RTL property via XML for paragraph")

def _flip_shape_position(shape, slide_width: int) -> None:
    """
    Flip text box position horizontally (mirror across slide center)

    Calculation:
    new_left = slide_width - (old_left + width)

    Example:
    - Slide width: 9144000 EMUs (10 inches)
    - Shape at left=1000000, width=3000000
    - New position: 9144000 - (1000000 + 3000000) = 5144000

    Args:
        shape: PowerPoint shape to flip
        slide_width: Width of the slide in EMUs (English Metric Units)
    """
    try:
        old_left = shape.left
        shape_width = shape.width

        # Calculate new left position (mirror)
        new_left = slide_width - (old_left + shape_width)

        # Clamp to valid range (prevent negative positions)
        new_left = max(0, new_left)
        new_left = min(new_left, slide_width - shape_width)

        # Apply new position
        shape.left = new_left

        logger.debug(f"Flipped shape position: {old_left} → {new_left}")

    except Exception as e:
        logger.warning(f"Could not flip shape position: {str(e)}")

def _set_arabic_font(shape) -> None:
    """
    Set Arabic-compatible font for all text runs

    Common Arabic fonts:
    - Arial (universally supported, good rendering)
    - Calibri (modern, clean)
    - Simplified Arabic
    - Traditional Arabic

    Args:
        shape: PowerPoint shape with text_frame
    """
    if not shape.has_text_frame:
        return

    font_name = Config.ARABIC_FONT

    text_frame = shape.text_frame

    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            try:
                # Set font name
                run.font.name = font_name

                # Ensure font applies to complex scripts (Arabic, Hebrew, etc.)
                # Access XML to set both Latin and Complex Script fonts
                rPr = run._element.rPr
                if rPr is not None:
                    # Set Latin font
                    latin_elem = rPr.find('.//a:latin', namespaces=NAMESPACES)
                    if latin_elem is not None:
                        latin_elem.set('typeface', font_name)
                    else:
                        # Create latin element
                        latin_elem = etree.SubElement(rPr, f"{{{NAMESPACES['a']}}}latin")
                        latin_elem.set('typeface', font_name)

                    # Set Complex Script font (for Arabic)
                    cs_elem = rPr.find('.//a:cs', namespaces=NAMESPACES)
                    if cs_elem is not None:
                        cs_elem.set('typeface', font_name)
                    else:
                        # Create cs element
                        cs_elem = etree.SubElement(rPr, f"{{{NAMESPACES['a']}}}cs")
                        cs_elem.set('typeface', font_name)

                logger.debug(f"Set Arabic font: {font_name}")

            except Exception as e:
                logger.warning(f"Could not set font for run: {str(e)}")

def reverse_bullet_order(text_frame) -> None:
    """
    Reverse the order of bullet points (OPTIONAL - not used in MVP)

    Some clients may want bullets visually reversed for RTL.
    For MVP, we skip this as it adds complexity and isn't always required.

    Args:
        text_frame: PowerPoint text_frame object
    """
    # Extract all paragraph data
    paragraphs_data = []

    for paragraph in text_frame.paragraphs:
        paragraphs_data.append({
            'text': paragraph.text,
            'level': paragraph.level,
            'font_name': paragraph.font.name if paragraph.font.name else None,
            'font_size': paragraph.font.size,
            'bold': paragraph.font.bold,
            'italic': paragraph.font.italic
        })

    # Clear existing paragraphs
    text_frame.clear()

    # Re-add in reverse order
    for data in reversed(paragraphs_data):
        p = text_frame.add_paragraph()
        p.text = data['text']
        p.level = data['level']

        if data['font_name']:
            p.font.name = data['font_name']
        if data['font_size']:
            p.font.size = data['font_size']
        if data['bold']:
            p.font.bold = data['bold']
        if data['italic']:
            p.font.italic = data['italic']

    logger.info("Reversed bullet order")

if __name__ == "__main__":
    # Test RTL conversion
    if len(sys.argv) > 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        flip_to_rtl_layout(input_file, output_file)
        print(f"RTL conversion complete: {output_file}")
    else:
        print("Usage: python rtl_converter.py <input.pptx> <output.pptx>")
