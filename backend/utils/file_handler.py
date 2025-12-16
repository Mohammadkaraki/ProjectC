"""
File handling utilities
"""
import os
import uuid
from pathlib import Path
from typing import Tuple
from werkzeug.utils import secure_filename
from config import Config
from utils.logger import setup_logger

logger = setup_logger(__name__)

def is_allowed_file(filename: str) -> bool:
    """
    Check if file extension is allowed

    Args:
        filename: Name of the file

    Returns:
        True if extension is allowed, False otherwise
    """
    return Path(filename).suffix.lower() in Config.ALLOWED_EXTENSIONS

def save_uploaded_file(file, filename: str) -> Tuple[str, str]:
    """
    Save uploaded file with unique ID

    Args:
        file: File object from request
        filename: Original filename

    Returns:
        Tuple of (file_id, file_path)
    """
    # Generate unique ID
    file_id = str(uuid.uuid4())

    # Secure the filename
    secure_name = secure_filename(filename)
    extension = Path(secure_name).suffix

    # Create new filename with UUID
    new_filename = f"{file_id}{extension}"
    file_path = Config.UPLOAD_FOLDER / new_filename

    # Save file
    file.save(str(file_path))
    logger.info(f"Saved uploaded file: {new_filename}")

    return file_id, str(file_path)

def get_output_path(file_id: str, suffix: str = "_translated") -> str:
    """
    Generate output file path

    Args:
        file_id: Unique file identifier
        suffix: Suffix to add to filename

    Returns:
        Output file path
    """
    output_filename = f"{file_id}{suffix}.pptx"
    output_path = Config.OUTPUT_FOLDER / output_filename
    return str(output_path)

def cleanup_temp_files(file_id: str):
    """
    Clean up temporary files

    Args:
        file_id: Unique file identifier
    """
    try:
        # Clean up upload file
        upload_file = Config.UPLOAD_FOLDER / f"{file_id}.pptx"
        if upload_file.exists():
            upload_file.unlink()
            logger.info(f"Cleaned up upload file: {file_id}.pptx")

        # Clean up output file (optional, keep for download)
        # output_file = Config.OUTPUT_FOLDER / f"{file_id}_translated.pptx"
        # if output_file.exists():
        #     output_file.unlink()

    except Exception as e:
        logger.warning(f"Error cleaning up files for {file_id}: {str(e)}")
