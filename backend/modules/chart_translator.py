"""
Chart Translator Module
Translates text within PowerPoint charts while keeping chart positions unchanged
"""
import os
import sys
import zipfile
import shutil
from lxml import etree
from typing import Dict, List, Any

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.logger import setup_logger
from modules.llm_translator import translate_with_openai

logger = setup_logger(__name__)

# XML namespaces for charts
CHART_NS = {
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

def translate_charts_in_pptx(input_path: str, output_path: str, source_lang: str = "English", target_lang: str = "Arabic"):
    """
    Translate all charts in a PowerPoint presentation
    
    Args:
        input_path: Path to input PPTX file
        output_path: Path to output PPTX file
        source_lang: Source language (default: English)
        target_lang: Target language (default: Arabic)
    """
    logger.info(f"Translating charts in: {input_path}")
    
    # Extract PPTX to temp directory
    temp_dir = input_path.replace('.pptx', '_chart_temp')
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    
    with zipfile.ZipFile(input_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    # Find all chart XML files
    charts_dir = os.path.join(temp_dir, 'ppt', 'charts')
    if not os.path.exists(charts_dir):
        logger.info("No charts found in presentation")
        _repack_pptx(temp_dir, output_path)
        shutil.rmtree(temp_dir)
        return
    
    chart_files = [f for f in os.listdir(charts_dir) if f.startswith('chart') and f.endswith('.xml')]
    logger.info(f"Found {len(chart_files)} chart file(s)")
    
    # Translate each chart
    for chart_file in chart_files:
        chart_path = os.path.join(charts_dir, chart_file)
        _translate_chart_file(chart_path, source_lang, target_lang)
    
    # Repack PPTX
    _repack_pptx(temp_dir, output_path)
    shutil.rmtree(temp_dir)
    
    logger.info(f"Chart translation complete: {output_path}")

def _translate_chart_file(chart_path: str, source_lang: str, target_lang: str):
    """Translate text in a single chart XML file"""
    logger.info(f"Translating chart: {os.path.basename(chart_path)}")
    
    # Parse XML
    tree = etree.parse(chart_path)
    root = tree.getroot()
    
    # Extract all text elements
    texts_to_translate = []
    text_elements = []
    
    # Find all <a:t> elements (text runs)
    for t_elem in root.findall('.//a:t', CHART_NS):
        if t_elem.text and t_elem.text.strip():
            texts_to_translate.append(t_elem.text.strip())
            text_elements.append(t_elem)
    
    if not texts_to_translate:
        logger.info(f"  No text found in chart")
        return
    
    logger.info(f"  Found {len(texts_to_translate)} text element(s) to translate")
    
    # Translate all texts using LLM
    # Create a simple structure for translation
    structure = {
        "elements": [
            {"element_id": f"text_{i}", "text": text, "type": "text_box"}
            for i, text in enumerate(texts_to_translate)
        ]
    }
    
    context = {}  # No context needed for chart translation
    
    translations = translate_with_openai(structure, context, source_lang, target_lang)
    
    # Replace text in XML
    for i, t_elem in enumerate(text_elements):
        element_id = f"text_{i}"
        if element_id in translations:
            translated_text = translations[element_id]
            logger.info(f"  '{t_elem.text.strip()[:30]}' -> '{translated_text[:30]}'")
            t_elem.text = translated_text
    
    # Save updated XML
    tree.write(chart_path, encoding='utf-8', xml_declaration=True, standalone=True)
    logger.info(f"  Updated chart XML")

def _repack_pptx(temp_dir: str, output_path: str):
    """Repack extracted PPTX directory into .pptx file"""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zipf.write(file_path, arcname)

if __name__ == "__main__":
    if len(sys.argv) > 2:
        translate_charts_in_pptx(sys.argv[1], sys.argv[2])
    else:
        print("Usage: python chart_translator.py <input.pptx> <output.pptx>")
