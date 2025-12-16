"""
Flask API Server for Slide Translator
Main entry point for the application
"""
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.exceptions import RequestEntityTooLarge
import os
import sys

# Add backend directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import Config
from utils.logger import setup_logger
from utils.file_handler import is_allowed_file, save_uploaded_file, get_output_path
from workflows.translate_all_slides import translate_all_slides
from modules.pdf_converter import convert_pdf_to_pptx, is_pdf_file

# Setup logger
logger = setup_logger(__name__)

# Initialize Flask app
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max file size
CORS(app)  # Enable CORS for frontend

# Validate configuration
try:
    Config.validate()
    logger.info("Configuration validated successfully")
except ValueError as e:
    logger.error(f"Configuration error: {str(e)}")
    sys.exit(1)

@app.route('/api/health', methods=['GET'])
def health_check():
    """
    Health check endpoint
    Returns: JSON status message
    """
    return jsonify({
        'status': 'ok',
        'service': 'Slide Translator API',
        'version': '1.0.0'
    })

@app.route('/api/translate-slide', methods=['POST'])
def translate_slide():
    """
    Main endpoint to translate PowerPoint slide from English to Arabic (RTL)

    Accepts: multipart/form-data with 'file' field containing .pptx file
    Returns: Translated .pptx file as download

    Workflow:
    1. Slide Ingestion & Parsing
    2. Context Understanding
    3. LLM-Based Translation
    4. RTL Layout Flip
    5. Text Replacement
    6. Output Generation
    """
    try:
        # Validate file in request
        if 'file' not in request.files:
            logger.warning("No file provided in request")
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']

        # Validate filename
        if file.filename == '':
            logger.warning("Empty filename")
            return jsonify({'error': 'No file selected'}), 400

        # Validate file extension
        if not is_allowed_file(file.filename):
            logger.warning(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'Only .pptx and .pdf files are allowed'}), 400

        logger.info(f"Received file: {file.filename}")

        # Save uploaded file
        file_id, input_path = save_uploaded_file(file, file.filename)
        output_path = get_output_path(file_id)

        # Check if PDF file - convert to PPTX first
        if is_pdf_file(file.filename):
            logger.info("PDF file detected. Converting to PowerPoint...")
            try:
                pptx_path = input_path.replace('.pdf', '_converted.pptx')
                input_path = convert_pdf_to_pptx(input_path, pptx_path)
                logger.info(f"PDF converted to PPTX: {pptx_path}")
            except Exception as e:
                logger.error(f"PDF conversion failed: {str(e)}")
                return jsonify({
                    'error': 'PDF conversion failed',
                    'message': 'Could not convert PDF to PowerPoint. Please ensure the PDF is not corrupted.'
                }), 400

        logger.info(f"Starting translation workflow for file_id: {file_id}")

        # Execute workflow (translates ALL slides)
        result_path = translate_all_slides(input_path, output_path)

        logger.info(f"Translation complete: {file_id}")

        # Return translated file
        return send_file(
            result_path,
            as_attachment=True,
            download_name='translated_slide.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        logger.error(f"Error during translation: {str(e)}", exc_info=True)
        return jsonify({
            'error': 'Translation failed',
            'message': str(e)
        }), 500

@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    """Handle file size limit exceeded"""
    logger.warning("File size limit exceeded")
    return jsonify({'error': 'File too large. Maximum size is 16 MB'}), 413

@app.errorhandler(404)
def not_found(e):
    """Handle 404 errors"""
    return jsonify({'error': 'Endpoint not found'}), 404

@app.errorhandler(500)
def internal_error(e):
    """Handle 500 errors"""
    logger.error(f"Internal server error: {str(e)}")
    return jsonify({'error': 'Internal server error'}), 500

if __name__ == '__main__':
    logger.info("Starting Slide Translator API server...")
    logger.info(f"Upload folder: {Config.UPLOAD_FOLDER}")
    logger.info(f"Output folder: {Config.OUTPUT_FOLDER}")

    app.run(
        host='0.0.0.0',
        port=5000,
        debug=Config.FLASK_DEBUG
    )
