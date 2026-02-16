from __future__ import annotations
import sys
from pathlib import Path

APP_NAME = "KP Generator"

# При сборке PyInstaller (--onefile) ресурсы лежат в sys._MEIPASS
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS) / "kp_generator"
else:
    BASE_DIR = Path(__file__).resolve().parent

ASSETS_DIR = BASE_DIR / "assets"

COMPANIES_JSON = ASSETS_DIR / "companies.json"

# Where to look for tesseract.exe (priority order)
# Для .exe: рядом с исполняемым файлом (sys.executable)
def _tesseract_candidates() -> list:
    candidates = [
        BASE_DIR / "tesseract" / "tesseract.exe",
        BASE_DIR / "tesseract.exe",
    ]
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates = [
            exe_dir / "tesseract" / "tesseract.exe",
            exe_dir / "tesseract.exe",
        ] + candidates
    return candidates

TESSERACT_CANDIDATES = _tesseract_candidates()

# Default output folder (создаётся при первой генерации)
def _default_output_dir() -> Path:
    preferred = Path.home() / "Documents" / "KP_Generator_Output"
    try:
        preferred.mkdir(parents=True, exist_ok=True)
        return preferred
    except OSError:
        if getattr(sys, "frozen", False):
            fallback = Path(sys.executable).resolve().parent / "output"
        else:
            fallback = BASE_DIR / "output"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback

DEFAULT_OUTPUT_DIR = _default_output_dir()

ROUNDING_STEPS = ["1", "10", "50", "100"]
