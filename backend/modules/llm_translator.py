"""
LLM Translator Module
Handles translation using OpenAI GPT-4 with context-aware prompts
"""
from openai import OpenAI
import json
from typing import Dict, List, Any
import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import Config
from utils.logger import setup_logger

logger = setup_logger(__name__)

# OpenAI client (initialized lazily)
_client = None

def get_openai_client():
    """Get or create OpenAI client instance"""
    global _client
    if _client is None:
        _client = OpenAI(api_key=Config.OPENAI_API_KEY)
    return _client

def translate_with_openai(
    slide_structure: Dict[str, Any],
    context_map: Dict[str, Any],
    source_lang: str = "English",
    target_lang: str = "Arabic"
) -> Dict[str, Any]:
    """
    Translate slide content using OpenAI with context awareness
    OPTIMIZED: Uses ONE API call per slide instead of multiple calls

    Args:
        slide_structure: Slide structure from slide_parser
        context_map: Context map from context_builder
        source_lang: Source language (default: English)
        target_lang: Target language (default: Arabic)

    Returns:
        Dictionary mapping element_id to translations:
        {
            "shape_0": "النهج الاستراتيجي لدينا",
            "shape_1": "النتائج الرئيسية من تحليل السوق",
            "shape_2": [
                "من المتوقع نمو الإيرادات بنسبة 15٪ سنويًا",
                "توسع الحصة السوقية في منطقة الشرق الأوسط"
            ]
        }
    """
    logger.info(f"Starting OPTIMIZED translation from {source_lang} to {target_lang}")

    try:
        # Build structured input for single API call
        elements_to_translate = {}

        for element in slide_structure["elements"]:
            element_id = element["element_id"]
            element_type = element["type"]

            if element_type == "title":
                text = element.get("text", "")
                if text:
                    elements_to_translate[element_id] = {
                        "type": "title",
                        "text": text
                    }
            elif element_type == "bullet_group":
                bullets = element.get("bullets", [])
                if bullets:
                    elements_to_translate[element_id] = {
                        "type": "bullets",
                        "items": [b["text"] for b in bullets]
                    }
            else:
                text = element.get("text", "")
                if text:
                    elements_to_translate[element_id] = {
                        "type": "text",
                        "text": text
                    }

        if not elements_to_translate:
            logger.info("No elements to translate")
            return {}

        # Make SINGLE API call to translate ALL elements
        logger.info(f"Translating {len(elements_to_translate)} elements in ONE API call")
        translations = _translate_slide_batch(
            elements_to_translate,
            source_lang,
            target_lang
        )

        logger.info(f"Translation complete. Translated {len(translations)} elements with 1 API call")
        return translations

    except Exception as e:
        logger.error(f"Translation error: {str(e)}", exc_info=True)
        raise

def _translate_slide_batch(
    elements: Dict[str, Any],
    source_lang: str,
    target_lang: str
) -> Dict[str, Any]:
    """
    Translate ALL elements from a slide in a SINGLE API call using JSON mode

    Args:
        elements: Dict mapping element_id to element data
            {
                "shape_0": {"type": "title", "text": "Our Strategy"},
                "shape_1": {"type": "text", "text": "Key findings"},
                "shape_2": {"type": "bullets", "items": ["Point 1", "Point 2"]}
            }
        source_lang: Source language
        target_lang: Target language

    Returns:
        Dict mapping element_id to translated content
            {
                "shape_0": "استراتيجيتنا",
                "shape_1": "النتائج الرئيسية",
                "shape_2": ["النقطة 1", "النقطة 2"]
            }
    """
    # Build structured prompt with all elements
    elements_list = []
    for element_id, element_data in elements.items():
        if element_data["type"] == "bullets":
            elements_list.append({
                "id": element_id,
                "type": "bullets",
                "content": element_data["items"]
            })
        else:
            elements_list.append({
                "id": element_id,
                "type": element_data["type"],
                "content": element_data["text"]
            })

    system_prompt = f"""You are a professional translator specializing in consulting and business presentations.
Translate from {source_lang} to {target_lang} while maintaining:
- Professional consulting tone and terminology
- Cultural appropriateness for MENA business audiences
- Concise, impactful language suitable for executive presentations
- Appropriate formality based on element type (titles are bold/impactful, body text is clear/professional)"""

    user_prompt = f"""Translate ALL elements from this slide from {source_lang} to {target_lang}.

Elements to translate:
{json.dumps(elements_list, ensure_ascii=False, indent=2)}

Return a JSON object with this exact structure:
{{
  "translations": {{
    "element_id_1": "translated text or array of translated items",
    "element_id_2": "translated text or array of translated items"
  }}
}}

Requirements:
- For "title" type: Translate as impactful, professional title
- For "text" type: Translate as clear, professional body text
- For "bullets" type: Return array of translated bullet points maintaining parallel structure
- Preserve the element IDs exactly as given
- Return ONLY the JSON object, no additional text"""

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

        # Extract translations
        translations = result.get("translations", {})

        # Validate we got all translations
        missing = set(elements.keys()) - set(translations.keys())
        if missing:
            logger.warning(f"Missing translations for: {missing}")
            # Fill in missing with placeholder
            for element_id in missing:
                translations[element_id] = "[Translation missing]"

        logger.info(f"Batch translated {len(translations)} elements successfully")
        return translations

    except Exception as e:
        logger.error(f"Batch translation error: {str(e)}")
        # Fallback: return error markers
        return {element_id: "[ERROR: Translation failed]" for element_id in elements.keys()}

def _translate_single_text(
    text: str,
    element_role: str,
    source_lang: str,
    target_lang: str,
    context: str
) -> str:
    """
    Translate a single text element

    Args:
        text: Text to translate
        element_role: Role of the element (title, header, etc.)
        source_lang: Source language
        target_lang: Target language
        context: Additional context for translation

    Returns:
        Translated text
    """
    system_prompt = f"""You are a professional translator specializing in consulting and business documents.
Translate from {source_lang} to {target_lang} while maintaining:
- Professional consulting tone and terminology
- Cultural appropriateness for MENA business audiences
- Concise, impactful language suitable for executive presentations"""

    user_prompt = f"""Translate this {element_role} from {source_lang} to {target_lang}:

Text: "{text}"

Context: {context}

Provide ONLY the {target_lang} translation, no explanations or additional text."""

    try:
        response = get_openai_client().chat.completions.create(
            model=Config.OPENAI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=Config.TRANSLATION_TEMPERATURE,
            max_tokens=Config.MAX_TOKENS
        )

        translation = response.choices[0].message.content.strip()
        logger.debug(f"Translated '{text[:50]}...' → '{translation[:50]}...'")
        return translation

    except Exception as e:
        logger.error(f"Error translating text: {str(e)}")
        raise

def _translate_bullet_list(
    bullets: List[str],
    source_lang: str,
    target_lang: str
) -> List[str]:
    """
    Translate a list of bullet points in a single batch

    Args:
        bullets: List of bullet point texts
        source_lang: Source language
        target_lang: Target language

    Returns:
        List of translated bullet points
    """
    system_prompt = f"""You are a professional translator specializing in consulting and business documents.
Translate from {source_lang} to {target_lang} while maintaining:
- Professional consulting tone
- Concise language (consulting slides should be brief)
- Parallel structure across bullets
- Cultural appropriateness for business audiences"""

    # Format bullets for translation
    numbered_bullets = "\n".join([f"{i+1}. {bullet}" for i, bullet in enumerate(bullets)])

    user_prompt = f"""Translate these consulting slide bullet points from {source_lang} to {target_lang}:

{numbered_bullets}

Requirements:
- Maintain professional consulting language
- Keep translations concise (suitable for slides)
- Preserve the hierarchy and structure

Return ONLY a JSON array of translated bullets, no additional text.
Format: ["translation 1", "translation 2", ...]"""

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

        # Parse JSON response
        try:
            # Try to parse as JSON object first
            result = json.loads(content)
            if isinstance(result, dict):
                # Extract array from dict (might be wrapped)
                translations = result.get("bullets", result.get("translations", list(result.values())[0] if result else []))
            else:
                translations = result
        except json.JSONDecodeError:
            # Fallback: split by newlines if JSON parsing fails
            logger.warning("Failed to parse JSON, using fallback splitting")
            translations = [line.strip().lstrip("0123456789.- ") for line in content.split("\n") if line.strip()]

        # Ensure we have the same number of translations
        if len(translations) != len(bullets):
            logger.warning(f"Translation count mismatch: {len(bullets)} original, {len(translations)} translated")
            # Pad or truncate as needed
            while len(translations) < len(bullets):
                translations.append("[Translation missing]")
            translations = translations[:len(bullets)]

        logger.debug(f"Translated {len(bullets)} bullets")
        return translations

    except Exception as e:
        logger.error(f"Error translating bullet list: {str(e)}")
        # Fallback: return original bullets with error marker
        return [f"[ERROR: {bullet}]" for bullet in bullets]

if __name__ == "__main__":
    # Test the translator
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        from modules.slide_parser import extract_slide_structure
        from modules.context_builder import build_context_map

        structure = extract_slide_structure(test_file)
        context = build_context_map(structure)
        translations = translate_with_openai(structure, context)

        print(json.dumps(translations, indent=2, ensure_ascii=False))
    else:
        print("Usage: python llm_translator.py <pptx_file>")
