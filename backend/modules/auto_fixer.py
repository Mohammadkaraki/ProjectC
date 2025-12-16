"""
Auto-Fixer Module
Automatically fixes common issues in translated presentations
"""
import zipfile
import tempfile
import shutil
import os
import sys
import re
from typing import Dict, List

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.logger import setup_logger

logger = setup_logger(__name__)

def auto_fix_presentation(input_path: str, output_path: str) -> Dict:
    """
    Automatically fix common issues in translated presentation

    Args:
        input_path: Path to input PPTX
        output_path: Path to save fixed PPTX

    Returns:
        Dictionary with fixes applied
    """
    logger.info(f"Auto-fixing presentation: {input_path}")

    # Extract PPTX
    temp_dir = tempfile.mkdtemp()
    extract_dir = os.path.join(temp_dir, 'pptx_extracted')

    with zipfile.ZipFile(input_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    # Apply fixes
    fixes_applied = {
        'alignment_fixes': 0,
        'font_fixes': 0,
        'margin_fixes': 0,
        'autofit_fixes': 0,
        'rtl_attribute_removals': 0
    }

    # Fix all slide XMLs
    slides_dir = os.path.join(extract_dir, 'ppt', 'slides')
    if os.path.exists(slides_dir):
        slide_files = [f for f in os.listdir(slides_dir) if f.endswith('.xml')]

        for slide_file in slide_files:
            slide_path = os.path.join(slides_dir, slide_file)
            slide_fixes = fix_slide_xml(slide_path)

            # Aggregate fixes
            for key, value in slide_fixes.items():
                fixes_applied[key] += value

    # Repack as PPTX
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, extract_dir)
                zipf.write(file_path, arcname)

    # Cleanup
    shutil.rmtree(temp_dir)

    logger.info(f"Auto-fix complete. Fixes applied: {sum(fixes_applied.values())}")

    return {
        'status': 'success',
        'total_fixes': sum(fixes_applied.values()),
        'details': fixes_applied
    }

def fix_slide_xml(xml_path: str) -> Dict:
    """
    Fix issues in a single slide XML file

    Args:
        xml_path: Path to slide XML file

    Returns:
        Dictionary with count of fixes applied
    """
    fixes = {
        'alignment_fixes': 0,
        'font_fixes': 0,
        'margin_fixes': 0,
        'autofit_fixes': 0,
        'rtl_attribute_removals': 0
    }

    # Read XML
    with open(xml_path, 'r', encoding='utf-8') as f:
        xml_content = f.read()

    original_content = xml_content

    # FIX 1: Standardize fonts to Arial
    # Replace common non-Arabic fonts with Arial
    fonts_to_replace = ['+mn-lt', 'Calibri', 'Times New Roman', 'Verdana', 'Tahoma']
    for font in fonts_to_replace:
        pattern = f'typeface="{font}"'
        if pattern in xml_content:
            xml_content = xml_content.replace(pattern, 'typeface="Arial"')
            fixes['font_fixes'] += xml_content.count('typeface="Arial"') - original_content.count('typeface="Arial"')

    # FIX 2: Fix alignment to RIGHT for Arabic text
    # Pattern: <a:pPr algn="l"> or <a:pPr algn="ctr"> -> should be algn="r"
    # But preserve images/shapes that shouldn't be aligned

    # Find all paragraph properties
    def fix_alignment(match):
        """Fix alignment in paragraph properties"""
        full_match = match.group(0)
        current_algn = match.group(1)

        # Skip if already right-aligned
        if current_algn == 'r':
            return full_match

        # Check if this paragraph contains Arabic text
        # (simplified check - in real scenario, would check content)
        fixes['alignment_fixes'] += 1
        return full_match.replace(f'algn="{current_algn}"', 'algn="r"')

    xml_content = re.sub(r'<a:pPr[^>]*algn="([^"]+)"[^>]*>', fix_alignment, xml_content)

    # FIX 3: Remove rtl="1" attribute (prevents text reversal)
    rtl_pattern = r'\s+rtl="1"'
    rtl_count = len(re.findall(rtl_pattern, xml_content))
    if rtl_count > 0:
        xml_content = re.sub(rtl_pattern, '', xml_content)
        fixes['rtl_attribute_removals'] = rtl_count
        logger.info(f"  Removed {rtl_count} rtl='1' attributes")

    # FIX 4: Optimize margins to 27432 EMUs (0.03 inches)
    optimal_margin = 27432

    def fix_margins(match):
        """Fix large margins in bodyPr"""
        full_match = match.group(0)

        # Get current margin values
        l_ins = match.group(1) if match.group(1) else '91440'
        r_ins = match.group(2) if match.group(2) else '91440'
        t_ins = match.group(3) if match.group(3) else '45720'
        b_ins = match.group(4) if match.group(4) else '45720'

        # Fix if larger than optimal
        l_ins_int = int(l_ins)
        r_ins_int = int(r_ins)

        if l_ins_int > optimal_margin or r_ins_int > optimal_margin:
            fixes['margin_fixes'] += 1
            # Replace with optimal margins
            new_bodypr = full_match
            if l_ins_int > optimal_margin:
                new_bodypr = new_bodypr.replace(f'lIns="{l_ins}"', f'lIns="{optimal_margin}"')
            if r_ins_int > optimal_margin:
                new_bodypr = new_bodypr.replace(f'rIns="{r_ins}"', f'rIns="{optimal_margin}"')
            return new_bodypr

        return full_match

    # Match bodyPr with margin attributes
    margin_pattern = r'<a:bodyPr[^>]*lIns="([^"]+)"[^>]*rIns="([^"]+)"[^>]*(?:tIns="([^"]+)"[^>]*)?(?:bIns="([^"]+)"[^>]*)?[^>]*>'
    xml_content = re.sub(margin_pattern, fix_margins, xml_content)

    # FIX 5: Add auto-fit if missing
    # Find bodyPr without normAutofit or spAutoFit
    def add_autofit(match):
        """Add normAutofit to bodyPr without auto-fit"""
        full_match = match.group(0)

        # Check if already has auto-fit
        if '<a:normAutofit' in full_match or '<a:spAutoFit' in full_match or '<a:noAutofit' in full_match:
            return full_match

        # Add normAutofit before closing tag
        if full_match.endswith('/>'):
            # Self-closing tag - add child element
            fixes['autofit_fixes'] += 1
            return full_match[:-2] + '><a:normAutofit/></a:bodyPr>'
        elif full_match.endswith('>'):
            # Opening tag - add child element
            fixes['autofit_fixes'] += 1
            return full_match + '<a:normAutofit/>'

        return full_match

    # Match bodyPr tags
    xml_content = re.sub(r'<a:bodyPr[^>]*>(?!</a:bodyPr>)', add_autofit, xml_content)

    # Write fixed XML
    if xml_content != original_content:
        with open(xml_path, 'w', encoding='utf-8') as f:
            f.write(xml_content)

    return fixes

if __name__ == "__main__":
    # Test auto-fixer
    if len(sys.argv) > 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2]

        result = auto_fix_presentation(input_file, output_file)

        print(f"\nAuto-Fix Results:")
        print(f"Status: {result['status']}")
        print(f"Total fixes: {result['total_fixes']}")
        print(f"\nDetails:")
        for fix_type, count in result['details'].items():
            if count > 0:
                print(f"  - {fix_type}: {count}")
    else:
        print("Usage: python auto_fixer.py <input.pptx> <output.pptx>")
