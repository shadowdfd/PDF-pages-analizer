"""Persistent high scores shared by the four bundled mini-games."""

import json
import os
from pathlib import Path
from typing import Dict


def _scores_path() -> Path:
    """Store scores in the user's application-data directory."""
    base = os.environ.get("APPDATA")
    if base:
        return Path(base) / "PDF pages PRO" / "highscores.json"
    return Path.home() / "PDF pages PRO" / "highscores.json"


def get_highscore(game: str) -> int:
    try:
        data = json.loads(_scores_path().read_text(encoding="utf-8"))
        return max(0, int(data.get(game, 0)))
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return 0


def save_highscore(game: str, score: int) -> int:
    """Persist a new record if higher; return the current record."""
    path = _scores_path()
    try:
        data: Dict[str, int] = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            data = {}
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        data = {}

    try:
        current = max(0, int(data.get(game, 0)))
    except (ValueError, TypeError):
        current = 0
    record = max(current, int(score), 0)
    data[game] = record

    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        # Scores remain usable for this session if storage is unavailable.
        pass
    return record
