"""
Translate ALL slides in a PowerPoint presentation
Simple wrapper that processes each slide
"""
import sys
import os
from pptx import Presentation
import shutil

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.slide_parser import extract_slide_structure
from modules.context_builder import build_context_map
from modules.llm_translator import translate_with_openai
from modules.rtl_converter import flip_to_rtl_layout
from modules.text_replacer import replace_text_in_slide
from modules.layout_translator import translate_slide_layouts
from config import Config
from utils.logger import setup_logger

logger = setup_logger(__name__)

def translate_all_slides(input_path: str, output_path: str):
    """Translate all slides in presentation"""

    logger.info("="*60)
    logger.info("Translating ALL slides")
    logger.info("="*60)

    # Get slide count
    prs = Presentation(input_path)
    slide_count = len(prs.slides)
    logger.info(f"Found {slide_count} slides")

    # Step 1-3: Parse, build context, translate ALL slides
    all_slides_data = []

    for slide_idx in range(slide_count):
        logger.info(f"\n[Slide {slide_idx+1}/{slide_count}] Parsing...")
        structure = extract_slide_structure(input_path, slide_idx)

        logger.info(f"[Slide {slide_idx+1}/{slide_count}] Building context...")
        context = build_context_map(structure)

        logger.info(f"[Slide {slide_idx+1}/{slide_count}] Translating...")
        translations = translate_with_openai(structure, context,
                                            Config.SOURCE_LANGUAGE,
                                            Config.TARGET_LANGUAGE)

        all_slides_data.append({
            'structure': structure,
            'translations': translations
        })

    # Step 4: Convert to RTL (all slides at once)
    logger.info("\nConverting to RTL layout...")
    rtl_temp = output_path.replace('.pptx', '_rtl_temp.pptx')
    flip_to_rtl_layout(input_path, rtl_temp)

    # Step 5: Replace text in ALL slides
    logger.info("\nReplacing text in all slides...")
    current_file = rtl_temp

    for slide_idx, slide_data in enumerate(all_slides_data):
        logger.info(f"[Slide {slide_idx+1}/{slide_count}] Replacing text...")

        # For last slide, output to final path
        out_file = output_path if slide_idx == slide_count-1 else current_file

        replace_text_in_slide(
            current_file,
            slide_data['translations'],
            slide_data['structure'],
            out_file,
            slide_idx
        )

    # Step 6: Translate layouts
    logger.info("\nTranslating layout backgrounds...")
    layout_out = output_path.replace('.pptx', '_with_layouts.pptx')
    translate_slide_layouts(output_path, layout_out)

    if os.path.exists(layout_out):
        os.remove(output_path)
        os.rename(layout_out, output_path)

    # Cleanup
    if os.path.exists(rtl_temp):
        os.remove(rtl_temp)

    logger.info("="*60)
    logger.info(f"SUCCESS! Translated {slide_count} slides")
    logger.info(f"Output: {output_path}")
    logger.info("="*60)

    return output_path

if __name__ == "__main__":
    if len(sys.argv) > 2:
        translate_all_slides(sys.argv[1], sys.argv[2])
    else:
        print("Usage: python translate_all_slides.py <input.pptx> <output.pptx>")
