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

    # LLM Provider Selection (openai or gemini)
    LLM_PROVIDER = os.getenv("LLM_PROVIDER", "gemini")  # Default to Gemini

    # OpenAI Configuration
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
    OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-3.5-turbo")

    # Gemini Configuration
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AIzaSyBSVMpPfkC54GMytJ-D1RtpKQbyvO1dm84")
    GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

    # Common Translation Settings
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
    ALLOWED_EXTENSIONS = {'.pptx'}

    # Arabic Font Settings
    ARABIC_FONT = "Arial"  # Universally supported

    @staticmethod
    def validate():
        """Validate required configuration"""
        provider = Config.LLM_PROVIDER.lower()

        if provider == "openai":
            if not Config.OPENAI_API_KEY:
                raise ValueError("OPENAI_API_KEY is required when using OpenAI provider. Please set it in .env file")
        elif provider == "gemini":
            if not Config.GEMINI_API_KEY:
                raise ValueError("GEMINI_API_KEY is required when using Gemini provider. Please set it in .env file")
        else:
            raise ValueError(f"Invalid LLM_PROVIDER: {provider}. Must be 'openai' or 'gemini'")

        return True
