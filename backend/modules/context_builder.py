"""
Context Builder Module
Builds context map to understand the role and significance of each text element
"""
from typing import Dict, List, Any
import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.logger import setup_logger

logger = setup_logger(__name__)

def build_context_map(slide_structure: Dict[str, Any]) -> Dict[str, Any]:
    """
    Build a context map that identifies the role of each text element

    Args:
        slide_structure: Output from extract_slide_structure()

    Returns:
        Dictionary mapping element_id to context information:
        {
            "shape_0": {
                "role": "slide_title",
                "context": "Main title of the consulting slide",
                "translation_priority": 1,
                "element_id": "shape_0"
            },
            "shape_1": {
                "role": "header",
                "context": "Key message or insight statement",
                "translation_priority": 2,
                "element_id": "shape_1"
            },
            ...
        }
    """
    logger.info("Building context map for slide elements")

    context_map = {}
    elements = slide_structure.get("elements", [])

    # Track what we've found
    has_title = False
    has_header = False

    for element in elements:
        element_id = element["element_id"]
        element_type = element["type"]

        # Determine role and context
        if element_type == "title" and not has_title:
            context_map[element_id] = {
                "role": "slide_title",
                "context": "Main title of the consulting slide - translate with professional, impactful language",
                "translation_priority": 1,
                "element_id": element_id,
                "type": element_type
            }
            has_title = True

        elif element_type == "header" and not has_header:
            context_map[element_id] = {
                "role": "header",
                "context": "Key message or insight statement - maintain executive summary tone",
                "translation_priority": 2,
                "element_id": element_id,
                "type": element_type
            }
            has_header = True

        elif element_type == "bullet_group":
            context_map[element_id] = {
                "role": "bullet_group",
                "context": "Supporting evidence or sub-points - maintain hierarchy and conciseness",
                "translation_priority": 3,
                "element_id": element_id,
                "type": element_type,
                "bullet_count": len(element.get("bullets", []))
            }

        elif element_type == "text_box":
            context_map[element_id] = {
                "role": "supporting_text",
                "context": "Additional information or notes",
                "translation_priority": 4,
                "element_id": element_id,
                "type": element_type
            }

        else:
            # Fallback for unknown types
            context_map[element_id] = {
                "role": "generic_text",
                "context": "Text element",
                "translation_priority": 5,
                "element_id": element_id,
                "type": element_type
            }

    logger.info(f"Built context map for {len(context_map)} elements")
    return context_map

def get_element_role(context_map: Dict[str, Any], element_id: str) -> str:
    """
    Get the role of a specific element

    Args:
        context_map: Context map from build_context_map()
        element_id: Element ID to look up

    Returns:
        Role string (e.g., "slide_title", "header", "bullet_group")
    """
    if element_id in context_map:
        return context_map[element_id]["role"]
    return "unknown"

def get_translation_instructions(context_map: Dict[str, Any], element_id: str) -> str:
    """
    Get specific translation instructions for an element

    Args:
        context_map: Context map from build_context_map()
        element_id: Element ID to look up

    Returns:
        Translation instruction string
    """
    if element_id in context_map:
        role = context_map[element_id]["role"]

        instructions = {
            "slide_title": "Translate as a professional, impactful title. Keep it concise and executive-level.",
            "header": "Translate as a key insight or finding. Maintain consulting tone and clarity.",
            "bullet_group": "Translate each bullet point concisely. Maintain hierarchy and professional language.",
            "supporting_text": "Translate supporting text while maintaining context and clarity.",
            "generic_text": "Translate accurately while preserving meaning."
        }

        return instructions.get(role, "Translate accurately")

    return "Translate accurately"

if __name__ == "__main__":
    # Test the context builder
    import json
    from modules.slide_parser import extract_slide_structure

    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        structure = extract_slide_structure(test_file)
        context = build_context_map(structure)
        print(json.dumps(context, indent=2))
    else:
        print("Usage: python context_builder.py <pptx_file>")
