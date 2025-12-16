"""
Slide Translator Workflow
Main workflow orchestrator that executes the complete translation pipeline
"""
import os
import sys
from pathlib import Path

# Add parent directory to path
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

class SlideTranslatorWorkflow:
    """
    Orchestrates the complete slide translation workflow

    Workflow Steps:
    1. Slide Ingestion & Parsing - Extract structure
    2. Context Understanding - Build context map
    3. LLM-Based Translation - Translate with OpenAI
    4. RTL Layout Flip - Convert to RTL layout
    5. Text Replacement - Insert translated text
    6. Output Generation - Save final slide
    """

    def __init__(self, input_path: str, output_path: str):
        """
        Initialize workflow

        Args:
            input_path: Path to input PowerPoint file
            output_path: Path to save translated output
        """
        self.input_path = input_path
        self.output_path = output_path

        # Workflow state
        self.state = {
            'current_step': None,
            'slide_count': 0,
            'slides_data': []  # List of {structure, context, translations} per slide
        }

        # Validate input file
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")

        logger.info(f"Initialized SlideTranslatorWorkflow")
        logger.info(f"Input: {input_path}")
        logger.info(f"Output: {output_path}")

    def execute(self) -> str:
        """
        Execute the complete workflow

        Returns:
            Path to final translated slide

        Raises:
            Exception if any step fails
        """
        logger.info("=" * 60)
        logger.info("Starting Slide Translator Workflow")
        logger.info("=" * 60)

        try:
            # Step 1: Parse slide structure
            self._step_1_parse_slide()

            # Step 2: Build context map
            self._step_2_build_context()

            # Step 3: Translate with LLM
            self._step_3_translate()

            # Step 4: Convert to RTL layout
            self._step_4_convert_rtl()

            # Step 5: Replace text
            self._step_5_replace_text()

            # Step 6: Translate layout backgrounds
            self._step_6_translate_layouts()

            # Step 7: Cleanup
            self._step_7_cleanup()

            logger.info("=" * 60)
            logger.info("Workflow Complete Successfully!")
            logger.info(f"Translated slide saved to: {self.output_path}")
            logger.info("=" * 60)

            return self.output_path

        except Exception as e:
            logger.error(f"Workflow failed at step: {self.state['current_step']}")
            logger.error(f"Error: {str(e)}", exc_info=True)
            raise

    def _step_1_parse_slide(self):
        """Step 1: Parse ALL slides structure"""
        self.state['current_step'] = 'parsing'
        logger.info("\n[STEP 1/7] Parsing slide structures...")

        # Get slide count
        from pptx import Presentation
        prs = Presentation(self.input_path)
        self.state['slide_count'] = len(prs.slides)

        logger.info(f"Found {self.state['slide_count']} slides to translate")

        # Parse each slide
        for slide_idx in range(self.state['slide_count']):
            logger.info(f"\nParsing slide {slide_idx + 1}/{self.state['slide_count']}...")
            slide_structure = extract_slide_structure(self.input_path, slide_idx)

            element_count = len(slide_structure['elements'])
            logger.info(f"  ✓ Extracted {element_count} elements")

            self.state['slides_data'].append({
                'structure': slide_structure,
                'context': None,
                'translations': None
            })

    def _step_2_build_context(self):
        """Step 2: Build context map for ALL slides"""
        self.state['current_step'] = 'context_building'
        logger.info("\n[STEP 2/7] Building context maps...")

        for slide_idx, slide_data in enumerate(self.state['slides_data']):
            logger.info(f"\nBuilding context for slide {slide_idx + 1}...")
            slide_data['context'] = build_context_map(slide_data['structure'])
            logger.info(f"  ✓ Built context for {len(slide_data['context'])} elements")

    def _step_3_translate(self):
        """Step 3: Translate ALL slides with LLM"""
        self.state['current_step'] = 'translating'
        logger.info("\n[STEP 3/7] Translating slides with OpenAI...")
        logger.info(f"  Model: {Config.OPENAI_MODEL}")
        logger.info(f"  Source: {Config.SOURCE_LANGUAGE} → Target: {Config.TARGET_LANGUAGE}")

        for slide_idx, slide_data in enumerate(self.state['slides_data']):
            logger.info(f"\nTranslating slide {slide_idx + 1}...")
            slide_data['translations'] = translate_with_openai(
                slide_data['structure'],
                slide_data['context'],
                source_lang=Config.SOURCE_LANGUAGE,
                target_lang=Config.TARGET_LANGUAGE
            )
            logger.info(f"  ✓ Translated {len(slide_data['translations'])} elements")

    def _step_4_convert_rtl(self):
        """Step 4: Convert to RTL layout (all slides at once)"""
        self.state['current_step'] = 'rtl_conversion'
        logger.info("\n[STEP 4/7] Converting ALL slides to RTL layout...")

        # Create temporary file for RTL conversion
        rtl_temp_path = self.output_path.replace('.pptx', '_rtl_temp.pptx')
        self.state['rtl_temp_path'] = rtl_temp_path

        flip_to_rtl_layout(self.input_path, rtl_temp_path)

        logger.info(f"✓ RTL conversion complete for all slides")

    def _step_5_replace_text(self):
        """Step 5: Replace text with translations in ALL slides"""
        self.state['current_step'] = 'text_replacement'
        logger.info("\n[STEP 5/7] Replacing text with translations...")

        # Start with RTL temp file
        current_file = self.state['rtl_temp_path']

        for slide_idx, slide_data in enumerate(self.state['slides_data']):
            logger.info(f"\nReplacing text in slide {slide_idx + 1}...")

            # For the last slide, output to final path
            output_file = self.output_path if slide_idx == len(self.state['slides_data']) - 1 else current_file

            replace_text_in_slide(
                current_file,
                slide_data['translations'],
                slide_data['structure'],
                output_file,
                slide_idx
            )

        logger.info(f"✓ Text replacement complete for all slides")

    def _step_6_translate_layouts(self):
        """Step 6: Translate slide layout backgrounds"""
        self.state['current_step'] = 'layout_translation'
        logger.info("\n[STEP 6/7] Translating slide layout backgrounds...")

        # Create temp output path for layout translation
        layout_output = self.output_path.replace('.pptx', '_with_layouts.pptx')

        translate_slide_layouts(self.output_path, layout_output)

        # Replace output with layout-translated version
        import os
        if os.path.exists(layout_output):
            os.remove(self.output_path)
            os.rename(layout_output, self.output_path)

        logger.info(f"✓ Layout backgrounds translated")

    def _step_7_cleanup(self):
        """Step 7: Cleanup temporary files"""
        self.state['current_step'] = 'cleanup'
        logger.info("\nCleaning up temporary files...")

        # Remove RTL temp file
        if self.state['rtl_temp_path'] and os.path.exists(self.state['rtl_temp_path']):
            try:
                os.remove(self.state['rtl_temp_path'])
                logger.info(f"✓ Removed temp file: {self.state['rtl_temp_path']}")
            except Exception as e:
                logger.warning(f"Could not remove temp file: {str(e)}")

        self.state['current_step'] = 'complete'

    def get_state(self) -> dict:
        """
        Get current workflow state

        Returns:
            Dictionary containing workflow state
        """
        return {
            'current_step': self.state['current_step'],
            'input_path': self.input_path,
            'output_path': self.output_path,
            'element_count': len(self.state['slide_structure']['elements']) if self.state['slide_structure'] else 0,
            'translation_count': len(self.state['translations']) if self.state['translations'] else 0
        }

if __name__ == "__main__":
    # Test workflow execution
    if len(sys.argv) > 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2]

        workflow = SlideTranslatorWorkflow(input_file, output_file)
        result = workflow.execute()

        print(f"\nSuccess! Translated slide: {result}")
    else:
        print("Usage: python slide_translator.py <input.pptx> <output.pptx>")
