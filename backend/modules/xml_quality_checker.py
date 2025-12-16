"""
XML Quality Checker Module
Deep inspection of PowerPoint XML to find and fix subtle issues
"""
import zipfile
import tempfile
import shutil
import os
import sys
import re
from typing import Dict, List
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import Config
from utils.logger import setup_logger
from modules.llm_translator import get_openai_client

logger = setup_logger(__name__)

def deep_xml_review(original_path: str, translated_path: str) -> Dict:
    """
    Perform deep XML-level review of translated presentation

    Args:
        original_path: Path to original PPTX
        translated_path: Path to translated PPTX

    Returns:
        Dictionary with detailed analysis and recommendations
    """
    logger.info("Starting deep XML quality check...")

    # Extract both presentations
    temp_dir = tempfile.mkdtemp()
    orig_dir = os.path.join(temp_dir, 'original')
    trans_dir = os.path.join(temp_dir, 'translated')

    with zipfile.ZipFile(original_path, 'r') as zip_ref:
        zip_ref.extractall(orig_dir)

    with zipfile.ZipFile(translated_path, 'r') as zip_ref:
        zip_ref.extractall(trans_dir)

    # Compare slide XMLs
    issues = []

    # Find all slide XML files
    orig_slides_dir = os.path.join(orig_dir, 'ppt', 'slides')
    trans_slides_dir = os.path.join(trans_dir, 'ppt', 'slides')

    if os.path.exists(orig_slides_dir) and os.path.exists(trans_slides_dir):
        orig_slides = sorted([f for f in os.listdir(orig_slides_dir) if f.endswith('.xml')])
        trans_slides = sorted([f for f in os.listdir(trans_slides_dir) if f.endswith('.xml')])

        for slide_file in orig_slides:
            if slide_file in trans_slides:
                orig_xml_path = os.path.join(orig_slides_dir, slide_file)
                trans_xml_path = os.path.join(trans_slides_dir, slide_file)

                slide_issues = analyze_slide_xml(orig_xml_path, trans_xml_path, slide_file)
                issues.extend(slide_issues)

    # Use AI to analyze issues and provide recommendations
    if issues:
        logger.info(f"Found {len(issues)} XML-level issues")
        recommendations = get_ai_recommendations(issues)
    else:
        logger.info("No XML-level issues found")
        recommendations = []

    # Cleanup
    shutil.rmtree(temp_dir)

    return {
        'status': 'complete',
        'issues_found': len(issues),
        'issues': issues,
        'recommendations': recommendations
    }

def analyze_slide_xml(orig_path: str, trans_path: str, slide_name: str) -> List[Dict]:
    """
    Analyze a single slide's XML for issues

    Args:
        orig_path: Path to original slide XML
        trans_path: Path to translated slide XML
        slide_name: Name of the slide file

    Returns:
        List of issues found
    """
    issues = []

    # Read XML content
    with open(orig_path, 'r', encoding='utf-8') as f:
        orig_xml = f.read()

    with open(trans_path, 'r', encoding='utf-8') as f:
        trans_xml = f.read()

    # Parse XML
    try:
        orig_tree = ET.fromstring(orig_xml)
        trans_tree = ET.fromstring(trans_xml)
    except ET.ParseError as e:
        logger.error(f"XML parse error in {slide_name}: {str(e)}")
        return issues

    # Define namespaces
    namespaces = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    # Check 1: Text overflow detection
    trans_text_elements = trans_tree.findall('.//a:t', namespaces)
    for idx, text_elem in enumerate(trans_text_elements):
        text = text_elem.text or ""
        if len(text) > 100:  # Long text might overflow
            # Check if parent has proper text fitting settings
            parent = text_elem
            text_body = None
            for _ in range(10):  # Search up to 10 levels
                parent = parent if hasattr(parent, 'getparent') else None
                if parent is None:
                    break
                if parent.tag.endswith('txBody'):
                    text_body = parent
                    break

            if text_body is not None:
                # Check for auto-fit settings
                bodypr = text_body.find('.//a:bodyPr', namespaces)
                if bodypr is not None:
                    wrap = bodypr.get('wrap', 'none')
                    if wrap == 'none':
                        issues.append({
                            'type': 'text_overflow_risk',
                            'severity': 'medium',
                            'slide': slide_name,
                            'description': f"Text element {idx+1} has no word wrap (length: {len(text)})",
                            'fix': 'Enable word wrap'
                        })

    # Check 2: Font consistency
    orig_fonts = set()
    trans_fonts = set()

    for latin_elem in orig_tree.findall('.//a:latin', namespaces):
        font = latin_elem.get('typeface')
        if font:
            orig_fonts.add(font)

    for latin_elem in trans_tree.findall('.//a:latin', namespaces):
        font = latin_elem.get('typeface')
        if font:
            trans_fonts.add(font)

    # Check if multiple fonts in translated (should be consistent Arial)
    if len(trans_fonts) > 2:  # Allow for 2 fonts (body + title)
        issues.append({
            'type': 'font_inconsistency',
            'severity': 'low',
            'slide': slide_name,
            'description': f"Multiple fonts detected in translated slide: {trans_fonts}",
            'fix': 'Standardize to Arial'
        })

    # Check 3: RTL alignment
    # Check if text is properly aligned to right
    algn_elements = trans_tree.findall('.//a:pPr', namespaces)
    for para_pr in algn_elements:
        algn = para_pr.get('algn', 'l')
        rtl = para_pr.get('rtl', '0')

        # For Arabic, alignment should be 'r' (right), and rtl should be '0' to prevent reversal
        if algn != 'r':
            issues.append({
                'type': 'alignment_issue',
                'severity': 'high',
                'slide': slide_name,
                'description': f"Text not aligned to right (algn={algn})",
                'fix': 'Set algn="r" for RTL text'
            })

        if rtl == '1':
            issues.append({
                'type': 'rtl_attribute_issue',
                'severity': 'high',
                'slide': slide_name,
                'description': 'rtl="1" will reverse Arabic text incorrectly',
                'fix': 'Remove rtl="1" attribute'
            })

    # Check 4: Margin settings
    bodypr_elements = trans_tree.findall('.//a:bodyPr', namespaces)
    for bodypr in bodypr_elements:
        # Get margin values (in EMUs)
        l_ins = int(bodypr.get('lIns', '91440'))  # Default 0.1 inch
        r_ins = int(bodypr.get('rIns', '91440'))
        t_ins = int(bodypr.get('tIns', '45720'))
        b_ins = int(bodypr.get('bIns', '45720'))

        # Optimal margins are 0.03 inches = 27432 EMUs
        optimal_margin = 27432

        if l_ins > optimal_margin * 1.5:
            issues.append({
                'type': 'large_margins',
                'severity': 'low',
                'slide': slide_name,
                'description': f"Left margin is {l_ins} EMUs (optimal: {optimal_margin})",
                'fix': f'Reduce margins to {optimal_margin} EMUs (0.03 inches)'
            })

    # Check 5: Auto-size settings
    for bodypr in bodypr_elements:
        # Check for proper auto-size
        has_normAutofit = bodypr.find('.//a:normAutofit', namespaces) is not None
        has_spAutoFit = bodypr.find('.//a:spAutoFit', namespaces) is not None

        if not has_normAutofit and not has_spAutoFit:
            issues.append({
                'type': 'no_autofit',
                'severity': 'medium',
                'slide': slide_name,
                'description': 'No auto-fit detected for text',
                'fix': 'Add <a:normAutofit/> for text-to-fit-shape'
            })

    return issues

def get_ai_recommendations(issues: List[Dict]) -> List[Dict]:
    """
    Use AI to analyze issues and provide actionable recommendations

    Args:
        issues: List of detected issues

    Returns:
        List of recommendations with priorities
    """
    if not issues:
        return []

    logger.info("Getting AI recommendations...")

    # Group issues by type
    issue_summary = {}
    for issue in issues:
        issue_type = issue['type']
        if issue_type not in issue_summary:
            issue_summary[issue_type] = []
        issue_summary[issue_type].append(issue)

    # Create prompt
    issue_text = "\n\n".join([
        f"**{issue_type}** ({len(items)} occurrences):\n" +
        "\n".join([f"  - {item['description']}" for item in items[:3]])  # Show max 3 examples
        for issue_type, items in issue_summary.items()
    ])

    prompt = f"""You are a PowerPoint XML expert reviewing a translated Arabic RTL presentation.

The following issues were detected:

{issue_text}

Context:
- Presentation was translated from English to Arabic
- RTL layout conversion was applied (horizontal mirroring)
- Arabic text uses Arial font
- Target: Professional business presentation

For each issue type, provide:
1. Severity assessment (critical/high/medium/low/ignore)
2. Professional impact (how it affects presentation quality)
3. Recommended action (specific fix to apply)
4. Priority (1-5, 1=most urgent)

Format as JSON:
{{
  "recommendations": [
    {{
      "issue_type": "...",
      "severity": "...",
      "impact": "...",
      "action": "...",
      "priority": 1
    }}
  ]
}}"""

    try:
        client = get_openai_client()

        response = client.chat.completions.create(
            model=Config.OPENAI_MODEL,
            messages=[
                {"role": "system", "content": "You are a PowerPoint XML expert specializing in professional Arabic RTL presentations."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            response_format={"type": "json_object"}
        )

        import json
        result = json.loads(response.choices[0].message.content)
        recommendations = result.get('recommendations', [])

        # Sort by priority
        recommendations.sort(key=lambda x: x.get('priority', 5))

        logger.info(f"Received {len(recommendations)} AI recommendations")
        return recommendations

    except Exception as e:
        logger.error(f"Error getting AI recommendations: {str(e)}")
        return []

if __name__ == "__main__":
    # Test deep XML review
    if len(sys.argv) > 2:
        original = sys.argv[1]
        translated = sys.argv[2]

        result = deep_xml_review(original, translated)

        print(f"\n{'='*60}")
        print("XML QUALITY CHECK RESULTS")
        print(f"{'='*60}\n")
        print(f"Status: {result['status']}")
        print(f"Issues found: {result['issues_found']}\n")

        if result['issues']:
            print("ISSUES DETECTED:")
            for idx, issue in enumerate(result['issues'][:10], 1):  # Show first 10
                print(f"\n{idx}. [{issue['severity'].upper()}] {issue['type']}")
                print(f"   Slide: {issue['slide']}")
                print(f"   Description: {issue['description']}")
                print(f"   Fix: {issue['fix']}")

        if result['recommendations']:
            print(f"\n{'='*60}")
            print("AI RECOMMENDATIONS:")
            print(f"{'='*60}\n")
            for idx, rec in enumerate(result['recommendations'], 1):
                print(f"{idx}. {rec['issue_type']} (Priority: {rec.get('priority', 'N/A')})")
                print(f"   Severity: {rec.get('severity', 'N/A')}")
                print(f"   Impact: {rec.get('impact', 'N/A')}")
                print(f"   Action: {rec.get('action', 'N/A')}\n")

    else:
        print("Usage: python xml_quality_checker.py <original.pptx> <translated.pptx>")
