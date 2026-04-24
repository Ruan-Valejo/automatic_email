from pathlib import Path
from dotenv import load_dotenv
import os

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
ATTACHMENTS_DIR = DATA_DIR / "attachments"

DATA_DIR.mkdir(exist_ok=True)
ATTACHMENTS_DIR.mkdir(exist_ok=True)

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///data/app.db")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")