from __future__ import annotations
import logging
from pathlib import Path
from datetime import datetime


def setup_file_logger(log_dir: Path) -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("kp_generator")
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    log_file = log_dir / f"kp_generator_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fmt = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger
