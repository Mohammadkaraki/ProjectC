# Slide Translator - Workflow Explanation

## Overview

This document provides a detailed, step-by-step explanation of how the Slide Translator workflow operates, from uploading an English PowerPoint slide to downloading a fully translated Arabic slide with RTL layout.

---

## Complete Workflow Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                       USER UPLOADS SLIDE                     │
│                     (English, LTR layout)                    │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 1: SLIDE INGESTION & PARSING                          │
│  Module: slide_parser.py                                     │
│                                                              │
│  Input:  sample_slide.pptx                                  │
│  Action: Extract all text elements and structure            │
│  Output: {                                                  │
│    "elements": [                                            │
│      {                                                      │
│        "element_id": "shape_0",                             │
│        "type": "title",                                     │
│        "text": "Market Expansion Strategy",                │
│        "position": {left: 100, top: 50, ...}               │
│      },                                                     │
│      {                                                      │
│        "element_id": "shape_1",                             │
│        "type": "header",                                    │
│        "text": "Key findings from MENA region analysis"    │
│      },                                                     │
│      {                                                      │
│        "element_id": "shape_2",                             │
│        "type": "bullet_group",                              │
│        "bullets": [                                         │
│          {text: "Revenue growth...", level: 0},            │
│          {text: "Market share...", level: 0},              │
│          {text: "Digital transformation...", level: 1}     │
│        ]                                                    │
│      }                                                      │
│    ]                                                        │
│  }                                                          │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 2: CONTEXT UNDERSTANDING                               │
│  Module: context_builder.py                                  │
│                                                              │
│  Input:  Slide structure from Step 1                        │
│  Action: Identify role of each element                      │
│  Output: {                                                  │
│    "shape_0": {                                             │
│      "role": "slide_title",                                 │
│      "context": "Main title - impactful language",         │
│      "translation_priority": 1                              │
│    },                                                       │
│    "shape_1": {                                             │
│      "role": "header",                                      │
│      "context": "Key message - executive tone",            │
│      "translation_priority": 2                              │
│    },                                                       │
│    "shape_2": {                                             │
│      "role": "bullet_group",                                │
│      "context": "Supporting evidence - concise",           │
│      "translation_priority": 3                              │
│    }                                                        │
│  }                                                          │
│                                                              │
│  Why: LLM translates better when it understands context     │
│       - Title ≠ Bullet (different tone, length, style)     │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 3: LLM-BASED TRANSLATION                               │
│  Module: llm_translator.py                                   │
│  Technology: OpenAI GPT-4                                    │
│                                                              │
│  Input:  Slide structure + Context map                      │
│  Action: Translate each element with context-aware prompts  │
│                                                              │
│  Example Translation Request:                               │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ System Prompt:                                       │  │
│  │ "You are a professional translator specializing in  │  │
│  │  consulting documents. Translate English to Arabic  │  │
│  │  with professional, executive-level tone."          │  │
│  │                                                      │  │
│  │ User Prompt:                                         │  │
│  │ "Translate this slide_title:                        │  │
│  │  'Market Expansion Strategy'                        │  │
│  │                                                      │  │
│  │  Context: Main title for executive presentation     │  │
│  │  Provide ONLY Arabic translation."                  │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  Output: {                                                  │
│    "shape_0": "استراتيجية التوسع في السوق",                 │
│    "shape_1": "النتائج الرئيسية من تحليل منطقة الشرق الأوسط",│
│    "shape_2": [                                             │
│      "من المتوقع نمو الإيرادات بنسبة 15٪ سنويًا",           │
│      "فرص توسع الحصة السوقية في الإمارات والسعودية",       │
│      "مبادرات التحول الرقمي تدفع الاعتماد"                 │
│    ]                                                        │
│  }                                                          │
│                                                              │
│  Key Features:                                              │
│  - Batch translation for consistency                        │
│  - Maintains hierarchy (title vs bullets)                   │
│  - Professional consulting terminology                      │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 4: RTL LAYOUT FLIP ⭐ MOST COMPLEX                     │
│  Module: rtl_converter.py                                    │
│  Technology: lxml (XML manipulation)                         │
│                                                              │
│  Challenge: PowerPoint doesn't have simple "flip to RTL"    │
│             python-pptx doesn't expose RTL properties       │
│                                                              │
│  Solution: Direct XML Manipulation                          │
│                                                              │
│  4A. Set Text Direction to RTL                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ # Access paragraph XML element                       │  │
│  │ p_element = paragraph._element                       │  │
│  │                                                      │  │
│  │ # Get or create <a:pPr> (paragraph properties)      │  │
│  │ pPr = etree.SubElement(                             │  │
│  │     p_element,                                       │  │
│  │     '{...}pPr'                                       │  │
│  │ )                                                    │  │
│  │                                                      │  │
│  │ # Set RTL attribute                                 │  │
│  │ pPr.set('rtl', '1')  # 1 = RTL, 0 = LTR            │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  4B. Set Text Alignment to RIGHT                            │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ from pptx.enum.text import PP_ALIGN                 │  │
│  │ paragraph.alignment = PP_ALIGN.RIGHT                │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  4C. Flip Shape Positions Horizontally                      │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ # Original: Shape at left=1000, width=3000          │  │
│  │ # Slide width=9144 (10 inches in EMUs)              │  │
│  │                                                      │  │
│  │ old_left = shape.left        # 1000                 │  │
│  │ width = shape.width          # 3000                 │  │
│  │ slide_width = prs.slide_width # 9144                │  │
│  │                                                      │  │
│  │ # Mirror position across slide center               │  │
│  │ new_left = slide_width - (old_left + width)        │  │
│  │          = 9144 - (1000 + 3000)                     │  │
│  │          = 5144                                      │  │
│  │                                                      │  │
│  │ shape.left = new_left                               │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  Before:                         After:                     │
│  ┌─────────────────────┐         ┌─────────────────────┐   │
│  │ Title    [shape]    │         │    [shape]    Title │   │
│  │ Bullet 1            │         │            Bullet 1 │   │
│  │ Bullet 2            │    →    │            Bullet 2 │   │
│  │ Bullet 3            │         │            Bullet 3 │   │
│  └─────────────────────┘         └─────────────────────┘   │
│  LTR (English)                   RTL (Arabic)               │
│                                                              │
│  4D. Set Arabic-Compatible Font                             │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ # Set font for both Latin and Complex Scripts       │  │
│  │ run.font.name = "Arial"                             │  │
│  │                                                      │  │
│  │ # Ensure Arabic rendering via XML                   │  │
│  │ latin_elem.set('typeface', 'Arial')                 │  │
│  │ cs_elem.set('typeface', 'Arial')  # Complex Script │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  Output: RTL-converted slide (still has English text)       │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 5: TEXT REPLACEMENT                                    │
│  Module: text_replacer.py                                    │
│                                                              │
│  Input:  RTL-converted slide + Translations                 │
│  Action: Replace English text with Arabic translations      │
│                                                              │
│  Process:                                                   │
│  1. Load RTL-converted slide                                │
│  2. Map element_id to shape objects                         │
│  3. For each element:                                       │
│     - If single text: Replace entire text                   │
│     - If bullets: Replace each bullet individually          │
│  4. Preserve RTL alignment (RIGHT)                          │
│  5. Save final output                                       │
│                                                              │
│  Example Replacement:                                       │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ # Single text element                                │  │
│  │ text_frame.clear()                                   │  │
│  │ paragraph = text_frame.add_paragraph()               │  │
│  │ paragraph.text = "استراتيجية التوسع في السوق"          │  │
│  │ paragraph.alignment = PP_ALIGN.RIGHT                 │  │
│  │                                                      │  │
│  │ # Bullet group                                       │  │
│  │ text_frame.clear()                                   │  │
│  │ for translation in translated_bullets:               │  │
│  │     p = text_frame.add_paragraph()                   │  │
│  │     p.text = translation                             │  │
│  │     p.level = original_level                         │  │
│  │     p.alignment = PP_ALIGN.RIGHT                     │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  Output: Fully translated Arabic slide with RTL layout      │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 6: OUTPUT GENERATION                                   │
│  Module: workflow orchestrator                               │
│                                                              │
│  Action:                                                    │
│  1. Save final slide to output path                         │
│  2. Clean up temporary files (RTL temp slide)               │
│  3. Return path to translated file                          │
│                                                              │
│  Output: translated_slide.pptx                              │
│  - Text in Arabic                                           │
│  - RTL layout (right-aligned, shapes flipped)               │
│  - Structure preserved (hierarchy, bullets)                 │
│  - Ready for consulting presentation                        │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│                  USER DOWNLOADS SLIDE                        │
│                (Arabic, RTL layout, ready to use)            │
└─────────────────────────────────────────────────────────────┘
```

---

## Detailed Step-by-Step Explanation

### STEP 1: Slide Ingestion & Parsing

**File:** `backend/modules/slide_parser.py`

**Purpose:** Extract all text content and structural information from the PowerPoint slide.

**How It Works:**
1. Load PowerPoint file using python-pptx: `Presentation(pptx_path)`
2. Access first slide: `slide = prs.slides[0]`
3. Iterate through all shapes on the slide
4. For each shape with text:
   - Identify type (title, header, bullet_group, text_box)
   - Extract text content
   - Capture position data (left, top, width, height)
   - For bullet groups: extract each bullet with its level (0=main, 1=sub-bullet)

**Key Code:**
```python
for shape in slide.shapes:
    if shape.has_text_frame:
        if shape.is_placeholder and shape.placeholder_format.type == 1:
            # This is a title
            element_type = "title"
        elif has multiple paragraphs:
            # This is a bullet group
            element_type = "bullet_group"
            bullets = extract_bullets(shape)
```

**Output Structure:**
```json
{
  "slide_index": 0,
  "elements": [
    {
      "element_id": "shape_0",
      "type": "title",
      "text": "Market Expansion Strategy",
      "shape_id": 2,
      "position": {"left": 635000, "top": 457200, ...}
    },
    {
      "element_id": "shape_2",
      "type": "bullet_group",
      "bullets": [
        {"text": "Revenue growth projected...", "level": 0, "index": 0},
        {"text": "Market share expansion...", "level": 0, "index": 1},
        {"text": "Digital transformation...", "level": 1, "index": 2}
      ]
    }
  ]
}
```

**Why This Step Matters:**
- Provides complete inventory of slide content
- Preserves hierarchy (which bullet is a sub-bullet?)
- Captures position data needed for RTL flipping later

---

### STEP 2: Context Understanding

**File:** `backend/modules/context_builder.py`

**Purpose:** Identify the role and significance of each text element to enable context-aware translation.

**How It Works:**
1. Receive slide structure from Step 1
2. For each element, determine its role:
   - **slide_title**: Main title (first title placeholder)
   - **header**: Key message or insight (first non-title text at top)
   - **bullet_group**: Supporting points or evidence
   - **supporting_text**: Additional notes
3. Assign translation priority (title = 1, header = 2, bullets = 3)
4. Generate context description for LLM

**Key Logic:**
```python
if element_type == "title" and not has_title:
    role = "slide_title"
    context = "Main title - translate with professional, impactful language"
    priority = 1

elif element_type == "bullet_group":
    role = "bullet_group"
    context = "Supporting evidence - maintain hierarchy and conciseness"
    priority = 3
```

**Output Structure:**
```json
{
  "shape_0": {
    "role": "slide_title",
    "context": "Main title of the consulting slide - translate with professional, impactful language",
    "translation_priority": 1
  },
  "shape_1": {
    "role": "header",
    "context": "Key message or insight statement - maintain executive summary tone",
    "translation_priority": 2
  },
  "shape_2": {
    "role": "bullet_group",
    "context": "Supporting evidence or sub-points - maintain hierarchy and conciseness",
    "translation_priority": 3,
    "bullet_count": 4
  }
}
```

**Why This Step Matters:**
- **Context-aware translation**: A title like "Strategic Approach" should be translated as a powerful headline, not a casual phrase
- **Tone consistency**: Headers need executive tone, bullets need concise language
- **Professional output**: Maintains consulting-grade language appropriate for each element type

**Example of Context Impact:**
- **Without context:** "Our Approach" → "نهجنا" (literal, casual)
- **With context (title):** "Our Approach" → "منهجنا الاستراتيجي" (strategic, professional)

---

### STEP 3: LLM-Based Translation

**File:** `backend/modules/llm_translator.py`

**Purpose:** Translate English text to Arabic using OpenAI GPT-4 with context-aware prompts.

**How It Works:**

**3A. Prompt Engineering**

System Prompt (sets AI behavior):
```
You are a professional translator specializing in consulting and business documents.
Translate from English to Arabic while maintaining:
- Professional consulting tone and terminology
- Cultural appropriateness for MENA business audiences
- Concise, impactful language suitable for executive presentations
```

User Prompt (specific translation request):
```
Translate this slide_title from English to Arabic:

Text: "Market Expansion Strategy"

Context: Main title of the consulting slide - translate with professional, impactful language

Provide ONLY the Arabic translation, no explanations or additional text.
```

**3B. Translation Strategy**

**Option 1: Element-by-Element Translation**
- Translate title separately
- Translate header separately
- Translate bullets separately
- **Pro:** Maximum context for each element
- **Con:** Multiple API calls (slower, more expensive)

**Option 2: Batch Translation (Implemented)**
- Translate all elements in one API call
- Use structured JSON format
- **Pro:** Single API call (faster, cheaper), consistent terminology
- **Con:** Requires careful prompt engineering

**Batch Translation Request:**
```json
{
  "title": "Market Expansion Strategy",
  "header": "Key findings from MENA region analysis",
  "bullets": [
    "Revenue growth projected at 15% annually through 2027",
    "Market share expansion opportunities in UAE and Saudi Arabia",
    "Digital transformation initiatives driving adoption",
    "Recommended investment: $2.5M for market entry phase"
  ]
}
```

**3C. API Call**
```python
from openai import OpenAI

client = OpenAI(api_key=Config.OPENAI_API_KEY)

response = client.chat.completions.create(
    model="gpt-4",  # or gpt-3.5-turbo for cost savings
    messages=[
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ],
    temperature=0.3,  # Lower = more consistent/deterministic
    max_tokens=2000,
    response_format={"type": "json_object"}  # Force JSON output
)

translations = json.loads(response.choices[0].message.content)
```

**3D. Translation Output**
```json
{
  "shape_0": "استراتيجية التوسع في السوق",
  "shape_1": "النتائج الرئيسية من تحليل منطقة الشرق الأوسط وشمال أفريقيا",
  "shape_2": [
    "من المتوقع نمو الإيرادات بنسبة 15٪ سنويًا حتى عام 2027",
    "فرص توسع الحصة السوقية في الإمارات العربية المتحدة والمملكة العربية السعودية",
    "مبادرات التحول الرقمي تدفع الاعتماد",
    "الاستثمار الموصى به: 2.5 مليون دولار لمرحلة دخول السوق"
  ]
}
```

**Why GPT-4 vs. Simple Translation:**
- **Context awareness:** Understands consulting jargon ("market entry phase" vs. literal words)
- **Cultural appropriateness:** Uses business-appropriate Arabic for MENA region
- **Tone consistency:** Maintains professional tone across elements
- **Hierarchical understanding:** Translates titles differently than supporting bullets

**Model Configuration:**
- **GPT-3.5-turbo (DEFAULT):** ~$0.002 per slide (fast, cost-effective, good quality)
- **GPT-4 (Optional):** ~$0.03 per slide (premium quality, better consulting tone)

The application uses GPT-3.5-turbo by default for optimal cost/performance balance.

---

### STEP 4: RTL Layout Flip ⭐ MOST COMPLEX

**File:** `backend/modules/rtl_converter.py`

**Purpose:** Convert PowerPoint slide layout from Left-to-Right (LTR) to Right-to-Left (RTL).

**The Challenge:**
PowerPoint does not provide a simple "flip to RTL" button via python-pptx API. The RTL text direction property is stored in the underlying XML, which python-pptx doesn't expose.

**The Solution: Direct XML Manipulation**

**4A. Set Text Direction via XML**

PowerPoint stores paragraph properties in XML namespace:
```xml
<a:p>
  <a:pPr rtl="1" algn="r"/>  <!-- rtl="1" = Right-to-Left -->
  <a:r>
    <a:t>Text here</a:t>
  </a:r>
</a:p>
```

Implementation:
```python
from lxml import etree

NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

def _set_rtl_via_xml(paragraph):
    # Get paragraph XML element
    p_element = paragraph._element

    # Get or create <a:pPr> (paragraph properties)
    pPr = p_element.find('.//a:pPr', namespaces=NAMESPACES)
    if pPr is None:
        pPr = etree.SubElement(p_element, f"{{{NAMESPACES['a']}}}pPr")

    # Set RTL attribute
    pPr.set('rtl', '1')  # 1 = RTL, 0 = LTR
```

**4B. Set Text Alignment to RIGHT**

python-pptx DOES support text alignment:
```python
from pptx.enum.text import PP_ALIGN

for paragraph in text_frame.paragraphs:
    paragraph.alignment = PP_ALIGN.RIGHT
```

**4C. Flip Shape Positions Horizontally**

Mirror text boxes across the slide's vertical center axis:

```python
def _flip_shape_position(shape, slide_width):
    """
    Flip shape position from left to right

    Calculation:
    new_left = slide_width - (old_left + width)

    Example:
    - Slide width: 9144000 EMUs (10 inches)
    - Shape at left=1000000, width=3000000
    - New position: 9144000 - (1000000 + 3000000) = 5144000
    """
    old_left = shape.left
    shape_width = shape.width

    # Calculate mirrored position
    new_left = slide_width - (old_left + shape_width)

    # Clamp to valid range (prevent negative positions)
    new_left = max(0, new_left)
    new_left = min(new_left, slide_width - shape_width)

    shape.left = new_left
```

**Visual Representation:**

```
BEFORE (LTR):
┌─────────────────────────────────────────┐
│ Title                                   │  ← left=500, width=8000
│ Header text                             │  ← left=500, width=8000
│ • Bullet 1                              │  ← left=800, width=7500
│ • Bullet 2                              │
└─────────────────────────────────────────┘
Slide width = 9144 EMUs

AFTER (RTL):
┌─────────────────────────────────────────┐
│                                   Title │  ← left=644 (9144-500-8000)
│                             Header text │  ← left=644
│                              • Bullet 1 │  ← left=844 (9144-800-7500)
│                              • Bullet 2 │
└─────────────────────────────────────────┘
```

**4D. Set Arabic-Compatible Font**

Ensure proper Arabic text rendering:
```python
def _set_arabic_font(shape, font_name="Arial"):
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name

            # Set font via XML for both Latin and Complex Scripts
            rPr = run._element.rPr

            # Latin font
            latin_elem = etree.SubElement(rPr, f"{{{NAMESPACES['a']}}}latin")
            latin_elem.set('typeface', font_name)

            # Complex Script font (Arabic, Hebrew, etc.)
            cs_elem = etree.SubElement(rPr, f"{{{NAMESPACES['a']}}}cs")
            cs_elem.set('typeface', font_name)
```

**Why Arial?**
- Universally available on all systems
- Excellent Arabic character support
- Professional appearance suitable for consulting

**Alternative fonts:**
- Calibri (modern, clean)
- Simplified Arabic (more traditional)
- Traditional Arabic (formal)

**Output of Step 4:**
A PowerPoint slide where:
- ✅ Text direction is RTL (via XML `rtl="1"`)
- ✅ Text is aligned to the right
- ✅ Text boxes are horizontally flipped
- ✅ Arabic font is set
- ⚠️ **Still contains English text** (translation happens in next step)

---

### STEP 5: Text Replacement

**File:** `backend/modules/text_replacer.py`

**Purpose:** Replace the original English text with Arabic translations in the RTL-converted slide.

**How It Works:**

**5A. Load RTL-Converted Slide**
```python
prs = Presentation(rtl_temp_path)  # From Step 4
slide = prs.slides[0]
```

**5B. Create Shape Mapping**
Map element_id to actual shape objects for quick lookup:
```python
shape_map = {}
for shape_idx, shape in enumerate(slide.shapes):
    element_id = f"shape_{shape_idx}"
    shape_map[element_id] = shape
```

**5C. Replace Text by Type**

**For Single Text Elements (title, header):**
```python
def _replace_single_text(shape, translation):
    text_frame = shape.text_frame

    # Clear existing paragraphs
    text_frame.clear()

    # Add new paragraph with translation
    paragraph = text_frame.add_paragraph()
    paragraph.text = translation

    # Maintain RTL alignment
    paragraph.alignment = PP_ALIGN.RIGHT
```

Example:
```python
# English: "Market Expansion Strategy"
# Arabic:  "استراتيجية التوسع في السوق"

_replace_single_text(shape, "استراتيجية التوسع في السوق")
```

**For Bullet Groups:**
```python
def _replace_bullets(shape, translations):
    text_frame = shape.text_frame

    # Store original bullet levels before clearing
    original_levels = [p.level for p in text_frame.paragraphs if p.text.strip()]

    # Clear existing bullets
    text_frame.clear()

    # Add translated bullets with original hierarchy
    for i, translation_text in enumerate(translations):
        paragraph = text_frame.add_paragraph()
        paragraph.text = translation_text
        paragraph.level = original_levels[i] if i < len(original_levels) else 0
        paragraph.alignment = PP_ALIGN.RIGHT
```

Example:
```python
# Original bullets (English, level 0):
# 1. "Revenue growth projected at 15% annually"
# 2. "Market share expansion opportunities"

# Translated bullets (Arabic, level 0 preserved):
# 1. "من المتوقع نمو الإيرادات بنسبة 15٪ سنويًا"
# 2. "فرص توسع الحصة السوقية"

_replace_bullets(shape, arabic_bullets)
```

**5D. Hierarchy Preservation**

Bullet levels are maintained:
```
Level 0 (Main bullet):
• من المتوقع نمو الإيرادات بنسبة 15٪ سنويًا

Level 1 (Sub-bullet, indented):
  • مبادرات التحول الرقمي تدفع الاعتماد
```

**5E. Save Final Output**
```python
prs.save(output_path)  # e.g., translated_slide.pptx
```

**Output of Step 5:**
A fully translated PowerPoint slide:
- ✅ Arabic text in all elements
- ✅ RTL layout (right-aligned, flipped positions)
- ✅ Hierarchy preserved (bullet levels intact)
- ✅ Professional fonts applied
- ✅ Ready for consulting presentation

---

### STEP 6: Output Generation & Cleanup

**File:** `backend/workflows/slide_translator.py`

**Purpose:** Finalize output and clean up temporary files.

**6A. Generate Final Output**
The slide from Step 5 is the final output, saved to:
```
tmp/outputs/{uuid}_translated.pptx
```

**6B. Cleanup Temporary Files**
Remove intermediate RTL conversion file:
```python
if os.path.exists(rtl_temp_path):
    os.remove(rtl_temp_path)
```

**6C. Return to User**
- **Web UI:** Download link provided
- **Command Line:** Path printed to console
- **API:** Binary file sent as HTTP response

**Final Output Verification:**
Open translated_slide.pptx in Microsoft PowerPoint:
- [ ] Text is in Arabic
- [ ] Text is aligned to the right
- [ ] Shapes are positioned on the right side
- [ ] Bullet hierarchy is preserved
- [ ] Professional consulting appearance

---

## Error Handling

**Throughout the workflow:**

1. **File Validation**
   - Check file exists
   - Verify .pptx format
   - Validate file size (<16 MB)

2. **API Error Handling**
   - OpenAI API failures (rate limits, invalid key)
   - Network errors
   - Timeout handling

3. **PowerPoint Errors**
   - Invalid slide structure
   - Missing text frames
   - XML parsing errors

4. **Graceful Degradation**
   - If XML manipulation fails → fallback to simple right-alignment
   - If font setting fails → continue with default
   - Detailed logging for debugging

**Example Error Flow:**
```python
try:
    workflow.execute()
except FileNotFoundError:
    return {"error": "Input file not found"}, 404
except openai.APIError as e:
    return {"error": f"Translation failed: {str(e)}"}, 500
except Exception as e:
    logger.error(f"Unexpected error: {str(e)}", exc_info=True)
    return {"error": "Internal server error"}, 500
```

---

## Performance Considerations

**Timing Breakdown (approximate):**
- Step 1 (Parsing): 0.5s
- Step 2 (Context): 0.1s
- Step 3 (Translation): 3-5s (depends on OpenAI API)
- Step 4 (RTL Conversion): 0.5s
- Step 5 (Text Replacement): 0.3s
- **Total:** ~5-7 seconds per slide

**Optimization Opportunities:**
- Use GPT-3.5-turbo instead of GPT-4 (faster, cheaper)
- Cache translations for repeated phrases
- Batch process multiple slides
- Async API calls for multi-slide decks

---

## Summary

The Slide Translator workflow transforms an English consulting slide into a professional Arabic RTL slide through 6 carefully orchestrated steps:

1. **Parse** → Extract structure
2. **Understand** → Identify context
3. **Translate** → AI-powered conversion with professional tone
4. **Flip RTL** → XML manipulation for true RTL layout
5. **Replace** → Insert Arabic text
6. **Output** → Deliver translated slide

**Key Innovations:**
- Context-aware AI translation (not just word-for-word)
- Direct XML manipulation for RTL (solving a complex technical challenge)
- Preservation of consulting-grade structure and hierarchy
- Production-ready error handling and logging

**Result:** A translated slide that looks and feels like it was created by a bilingual consultant, not a machine translation tool.

---

**For more details, see:**
- [README.md](../README.md) - Setup and usage instructions
- [ARCHITECTURE.md](ARCHITECTURE.md) - System design and technical decisions
- Backend code in `backend/` directory for implementation details
