"""
Configuration module for Slide Translator application
"""
import os
from dotenv import load_dotenv
from pathlib import Path

# Load environment variables
load_dotenv()

class Config:
    """Application configuration"""

    # OpenAI Configuration
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
    OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-3.5-turbo")
    TRANSLATION_TEMPERATURE = float(os.getenv("TRANSLATION_TEMPERATURE", "0.3"))
    MAX_TOKENS = int(os.getenv("MAX_TOKENS", "2000"))

    # Flask Configuration
    FLASK_ENV = os.getenv("FLASK_ENV", "development")
    FLASK_DEBUG = os.getenv("FLASK_DEBUG", "True").lower() == "true"

    # File Paths
    BASE_DIR = Path(__file__).parent.parent
    UPLOAD_FOLDER = BASE_DIR / "tmp" / "uploads"
    OUTPUT_FOLDER = BASE_DIR / "tmp" / "outputs"

    # Ensure directories exist
    UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    # Logging
    LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")

    # Translation Settings
    SOURCE_LANGUAGE = "English"
    TARGET_LANGUAGE = "Arabic"

    # Supported file types
    ALLOWED_EXTENSIONS = {'.pptx', '.pdf'}

    # Arabic Font Settings
    ARABIC_FONT = "Arial"  # Universally supported

    @staticmethod
    def validate():
        """Validate required configuration"""
        if not Config.OPENAI_API_KEY:
            raise ValueError("OPENAI_API_KEY is required. Please set it in .env file")

        return True
