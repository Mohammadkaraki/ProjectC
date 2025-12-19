"""
Layout Translator Module
Translates text in slide layouts (background/template text)
"""
import zipfile
import tempfile
import shutil
import os
import re
from typing import Dict, List
import sys

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import Config
from utils.logger import setup_logger
from modules.llm_translator import _translate_single_text

logger = setup_logger(__name__)

def _batch_translate_layout_texts(
    texts: List[str],
    source_lang: str,
    target_lang: str
) -> List[str]:
    """
    Batch translate multiple layout texts in ONE API call

    Args:
        texts: List of text strings to translate
        source_lang: Source language
        target_lang: Target language

    Returns:
        List of translated strings in same order
    """
    from modules.llm_translator import get_openai_client
    import json

    if not texts:
        return []

    # Build prompt for batch translation
    numbered_texts = "\n".join([f"{i+1}. {text}" for i, text in enumerate(texts)])

    system_prompt = f"""You are a professional translator for PowerPoint slide layouts and templates.
Translate from {source_lang} to {target_lang} while maintaining professional business tone."""

    user_prompt = f"""Translate these layout/background texts from {source_lang} to {target_lang}:

{numbered_texts}

Return ONLY a JSON object with this structure:
{{
  "translations": ["translation 1", "translation 2", ...]
}}

Preserve the order. Return ONLY the JSON, no additional text."""

    try:
        response = get_openai_client().chat.completions.create(
            model=Config.OPENAI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=Config.TRANSLATION_TEMPERATURE,
            max_tokens=Config.MAX_TOKENS,
            response_format={"type": "json_object"}
        )

        content = response.choices[0].message.content.strip()
        result = json.loads(content)
        translations = result.get("translations", [])

        # Ensure we have the same number
        if len(translations) != len(texts):
            logger.warning(f"Translation count mismatch: expected {len(texts)}, got {len(translations)}")
            # Pad with originals if missing
            while len(translations) < len(texts):
                translations.append(texts[len(translations)])

        return translations[:len(texts)]

    except Exception as e:
        logger.error(f"Batch layout translation error: {str(e)}")
        # Fallback: return originals
        return texts

def translate_slide_layouts(pptx_path: str, output_path: str) -> None:
    """
    Translate text in slide layouts (background graphics)

    Args:
        pptx_path: Path to input PPTX file
        output_path: Path to save output PPTX file with translated layouts
    """
    logger.info(f"Translating slide layouts in: {pptx_path}")

    try:
        # Create temp directory
        temp_dir = tempfile.mkdtemp()
        extract_dir = os.path.join(temp_dir, 'pptx_extracted')

        # Extract PPTX as ZIP
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        # Find and translate all slide layouts
        layout_dir = os.path.join(extract_dir, 'ppt', 'slideLayouts')

        if not os.path.exists(layout_dir):
            logger.warning("No slide layouts found")
            shutil.copy(pptx_path, output_path)
            return

        layout_files = [f for f in os.listdir(layout_dir) if f.endswith('.xml')]
        logger.info(f"Found {len(layout_files)} layout files")

        # Translate each layout
        for layout_file in layout_files:
            layout_path = os.path.join(layout_dir, layout_file)
            _translate_layout_file(layout_path)

        # Repack as PPTX
        _repack_pptx(extract_dir, output_path)

        # Cleanup
        shutil.rmtree(temp_dir)

        logger.info(f"Layout translation complete: {output_path}")

    except Exception as e:
        logger.error(f"Error translating layouts: {str(e)}", exc_info=True)
        # Fallback: copy original if translation fails
        shutil.copy(pptx_path, output_path)
        raise

def _translate_layout_file(layout_xml_path: str) -> None:
    """
    Translate text elements in a layout XML file
    OPTIMIZED: Batches all layout text into ONE API call

    Args:
        layout_xml_path: Path to slideLayout XML file
    """
    logger.info(f"Translating layout: {os.path.basename(layout_xml_path)}")

    # Read XML
    with open(layout_xml_path, 'r', encoding='utf-8') as f:
        xml_content = f.read()

    # Find all text elements with <a:t>text</a:t>
    pattern = r'<a:t>([^<]+)</a:t>'
    matches = list(re.finditer(pattern, xml_content))

    # Collect all texts to translate
    texts_to_translate = []
    text_map = {}  # Map index to original text with spaces

    for idx, match in enumerate(matches):
        original_with_spaces = match.group(1)
        original_text = original_with_spaces.strip()

        # Skip empty, numbers, or very short text
        if not original_text or len(original_text) < 2:
            continue

        # Skip if already translated (contains Arabic)
        if any('\u0600' <= c <= '\u06FF' for c in original_text):
            continue

        texts_to_translate.append(original_text)
        text_map[len(texts_to_translate) - 1] = original_with_spaces

    if not texts_to_translate:
        logger.info("  No text to translate in this layout")
        return

    # BATCH TRANSLATE: All texts in ONE API call
    logger.info(f"  Translating {len(texts_to_translate)} text elements in ONE API call...")
    try:
        translated_list = _batch_translate_layout_texts(
            texts_to_translate,
            Config.SOURCE_LANGUAGE,
            Config.TARGET_LANGUAGE
        )

        # Build translations dict
        translations = {}
        for idx, translated in enumerate(translated_list):
            original_with_spaces = text_map[idx]
            translations[original_with_spaces] = translated
            logger.info(f"  ✓ '{original_with_spaces.strip()[:50]}...' -> '{translated[:50]}...'")

    except Exception as e:
        logger.error(f"  Batch translation failed: {str(e)}")
        return

    # Replace in XML
    for original, translated in translations.items():
        # Escape XML special characters
        original_escaped = original.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        translated_escaped = translated.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        xml_content = xml_content.replace(
            f'<a:t>{original_escaped}</a:t>',
            f'<a:t>{translated_escaped}</a:t>'
        )

    # Write back
    with open(layout_xml_path, 'w', encoding='utf-8') as f:
        f.write(xml_content)

    logger.info(f"  Translated {len(translations)} text elements in layout")

def _flip_layout_shapes_rtl(xml_content: str) -> str:
    """
    Flip shape positions in layout XML to RTL (mirror horizontally)

    Args:
        xml_content: Layout XML content

    Returns:
        Modified XML with flipped shape positions
    """
    logger.info("  Flipping layout shapes to RTL...")

    # Standard slide width in EMUs (12192000 for 16:9, 9144000 for 4:3)
    # We'll extract it from the XML or use default
    slide_width_match = re.search(r'<p:sldSz\s+cx="(\d+)"', xml_content)
    slide_width = int(slide_width_match.group(1)) if slide_width_match else 12192000

    # Pattern to find shape positions: <a:off x="..." y="..."/>
    # and shape extents: <a:ext cx="..." cy="..."/>

    def flip_position(match):
        """Flip a single shape position"""
        x = int(match.group(1))
        cx = int(match.group(2))  # width

        # Calculate flipped position: new_x = slide_width - (old_x + width)
        new_x = slide_width - (x + cx)

        return f'x="{new_x}"'

    # Find all <p:sp> (shape) elements and flip their positions
    # Pattern matches: <a:off x="NUMBER" followed later by <a:ext cx="NUMBER"

    # More robust approach: find x position and width together
    pattern = r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)"'

    def replace_offset(match):
        x = int(match.group(1))
        y = match.group(2)  # Keep Y unchanged
        cx = int(match.group(3))

        # Flip X position
        new_x = slide_width - (x + cx)

        return f'<a:off x="{new_x}" y="{y}"/><a:ext cx="{cx}"'

    xml_content = re.sub(pattern, replace_offset, xml_content)

    logger.info(f"  Flipped layout shapes (slide width: {slide_width} EMUs)")

    return xml_content

def _repack_pptx(extract_dir: str, output_path: str) -> None:
    """
    Repack extracted directory as PPTX file

    Args:
        extract_dir: Directory containing extracted PPTX contents
        output_path: Path to save repacked PPTX
    """
    logger.info("Repacking PPTX...")

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, extract_dir)
                zipf.write(file_path, arcname)

    logger.info(f"Repacked PPTX: {output_path}")

if __name__ == "__main__":
    # Test layout translation
    if len(sys.argv) > 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2]

        translate_slide_layouts(input_file, output_file)
        print(f"Layout translation complete: {output_file}")
    else:
        print("Usage: python layout_translator.py <input.pptx> <output.pptx>")
