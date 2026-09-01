"""Bộ nhớ đệm kết quả phân loại email, tránh gọi lại Gemini cho thư đã xử lý."""
from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

from config import CACHE_DIR

logger = logging.getLogger(__name__)


class ClassificationCache:
    def __init__(self, cache_dir: Optional[Path] = None, max_entries: int = 2000) -> None:
        self.path = (Path(cache_dir) if cache_dir is not None else CACHE_DIR) / "classifications.json"
        self.max_entries = max_entries
        self._data: Dict[str, Dict] = self._load()

    def _load(self) -> Dict[str, Dict]:
        if not self.path.exists():
            return {}
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception as e:
            logger.warning("Không đọc được cache phân loại: %s", e)
            return {}

    def _save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False)
        except Exception as e:
            logger.warning("Không ghi được cache phân loại: %s", e)

    def get_many(self, entry_ids: List[str]) -> Dict[str, Dict]:
        return {eid: self._data[eid] for eid in entry_ids if eid in self._data}

    def put_many(self, mapping: Dict[str, Dict]) -> None:
        if not mapping:
            return
        now = datetime.now(timezone.utc).isoformat()
        for entry_id, value in mapping.items():
            self._data[entry_id] = {**value, "cached_at": now}
        self._evict()
        self._save()

    def _evict(self) -> None:
        """Giữ lại max_entries mục mới nhất theo cached_at."""
        if len(self._data) <= self.max_entries:
            return
        ordered = sorted(self._data.items(),
                         key=lambda kv: kv[1].get("cached_at", ""), reverse=True)
        self._data = dict(ordered[:self.max_entries])
