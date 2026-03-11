"""Persist sync state to avoid re-writing unchanged emails."""

from __future__ import annotations

import json
from pathlib import Path


class SyncState:
    """Tracks which entry IDs have already been synced."""

    def __init__(self, state_file: Path) -> None:
        self._path = state_file
        self._seen: set[str] = set()
        self._load()

    def _load(self) -> None:
        if self._path.exists():
            try:
                data = json.loads(self._path.read_text(encoding="utf-8"))
                self._seen = set(data.get("synced", []))
            except Exception:
                self._seen = set()

    def save(self) -> None:
        self._path.write_text(
            json.dumps({"synced": sorted(self._seen)}, indent=2),
            encoding="utf-8",
        )

    @property
    def seen(self) -> set[str]:
        return self._seen

    def is_synced(self, entry_id: str) -> bool:
        return entry_id in self._seen

    def mark_synced(self, entry_id: str) -> None:
        self._seen.add(entry_id)

    @property
    def count(self) -> int:
        return len(self._seen)
