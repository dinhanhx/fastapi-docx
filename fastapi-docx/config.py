"""Configuration loaded from environment variables via .env file."""

import os

from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Server Configuration
PORT = int(os.getenv("PORT", 9700))
WORKERS = int(os.getenv("WORKERS", 1))
RELOAD = os.getenv("RELOAD", "false").lower() in ("true", "1", "yes")
MAX_UPLOAD_BYTES = int(os.getenv("MAX_UPLOAD_BYTES", 5 * 1024 * 1024))  # 5 MB default
CONVERSION_TIMEOUT = float(os.getenv("CONVERSION_TIMEOUT", 120))  # seconds per conversion
