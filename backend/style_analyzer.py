"""Lưu/đọc hồ sơ văn phong dạng JSON.

Quy trình thống kê thật nằm ở backend/style_stats.py (đếm bằng Python) và được
EmailAssistantAPI.export_sent_emails gọi — module này chỉ lo phần ghi đĩa.
"""
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

logger = logging.getLogger(__name__)


class StyleAnalyzer:
    def __init__(self, data_dir: str):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def save_profile(self, profile: Dict, profile_name: str = 'default') -> Dict:
        """Lưu hồ sơ đã được tổng hợp từ nhiều lô phân tích AI."""
        profile = dict(profile or {})
        profile['name'] = profile_name
        profile['analyzed_at'] = datetime.now().astimezone().isoformat()
        filepath = self.data_dir / f"{profile_name}.json"
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(profile, f, ensure_ascii=False, indent=2)
        logger.info("Đã lưu hồ sơ văn phong: %s", filepath)
        return profile

    def load_profile(self, profile_name: str = 'default') -> Optional[Dict]:
        filepath = self.data_dir / f"{profile_name}.json"
        if not filepath.exists():
            return None
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (OSError, UnicodeError, json.JSONDecodeError) as e:
            logger.error("Không đọc được hồ sơ %s: %s", profile_name, e)
            return None
