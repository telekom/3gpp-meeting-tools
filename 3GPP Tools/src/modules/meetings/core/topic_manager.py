# --- File: src/modules/meetings/core/topic_manager.py ---
"""
Persistent Meta-Topic & Multi-WG Topic Manager.

Provides data structures, JSON persistence, and fast in-memory rule matching
to correlate 3GPP contributions and decisions across Working Groups (SA, RAN, CT)
and meeting cycles.
"""

import copy
import fnmatch
import json
import logging
import os
import re
import uuid
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from core.utils.paths import get_project_root, to_extended_path

logger = logging.getLogger(__name__)

DEFAULT_CONFIG_PATH = Path("config/meta_topics.json")
USER_CONFIG_PATH = Path("config/meta_topics_user.json")


def _extract_base_tokens(text: str) -> Set[str]:
    """
    Extracts normalized tokens from Work Item acronyms.
    E.g., 'AmbientIoT_Ph2-CT' -> {'ambientiot_ph2', 'ambientiot'}.
    """
    if not text:
        return set()
    cleaned = re.sub(r'[\r\n\t]+', ' ', str(text)).strip().lower()
    tokens = set()
    for raw in re.split(r'[,;\s/&|]+', cleaned):
        t = raw.strip()
        if not t or len(t) < 2:
            continue
        # Strip trailing WG tags like '-ct', '-ran', '-sa'
        base = re.sub(r'-(?:ct|ran|sa|ct1|ct3|ct4|ran1|ran2|ran3|sa2|sa3)$', '', t)
        tokens.add(base)
        # Also strip leading FS_ / study prefix
        if base.startswith("fs_"):
            tokens.add(base[3:])
        # Strip Ph2 / Phase suffixes for root matching
        root = re.sub(r'_ph\d+$', '', base)
        if root != base:
            tokens.add(root)
            if root.startswith("fs_"):
                tokens.add(root[3:])
    return tokens


@dataclass
class MetaTopicRules:
    wi_tokens: List[str] = field(default_factory=list)
    ai_patterns: Dict[str, List[str]] = field(default_factory=dict)
    keywords: List[str] = field(default_factory=list)
    pinned_tdocs: List[str] = field(default_factory=list)
    excluded_tdocs: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "wi_tokens": list(self.wi_tokens),
            "ai_patterns": {k: list(v) for k, v in self.ai_patterns.items()},
            "keywords": list(self.keywords),
            "pinned_tdocs": [t.upper().strip() for t in self.pinned_tdocs if t.strip()],
            "excluded_tdocs": [t.upper().strip() for t in self.excluded_tdocs if t.strip()],
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "MetaTopicRules":
        if not data:
            return cls()
        return cls(
            wi_tokens=[str(w).strip().lower() for w in data.get("wi_tokens", []) if str(w).strip()],
            ai_patterns={
                str(k).upper().strip(): [str(p).strip().lower() for p in v if str(p).strip()]
                for k, v in data.get("ai_patterns", {}).items()
            },
            keywords=[str(k).strip().lower() for k in data.get("keywords", []) if str(k).strip()],
            pinned_tdocs=[str(t).upper().strip() for t in data.get("pinned_tdocs", []) if str(t).strip()],
            excluded_tdocs=[str(t).upper().strip() for t in data.get("excluded_tdocs", []) if str(t).strip()],
        )


@dataclass
class MetaTopicDefinition:
    id: str
    name: str
    description: str = ""
    color: str = "#10B981"
    target_wgs: List[str] = field(default_factory=lambda: ["SA2", "RAN2", "RAN3", "CT1"])
    cycle_mode: str = "past_n_cycles"  # "past_n_cycles", "cycle_range", "date_range"
    past_n_cycles: int = 4
    from_cycle: str = "2025-Q1"
    to_cycle: str = "2026-Q1"
    date_from: str = ""
    date_to: str = ""
    rules: MetaTopicRules = field(default_factory=MetaTopicRules)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "color": self.color,
            "target_wgs": [w.upper().strip() for w in self.target_wgs if w.strip()],
            "cycle_mode": self.cycle_mode,
            "past_n_cycles": int(self.past_n_cycles or 4),
            "from_cycle": self.from_cycle,
            "to_cycle": self.to_cycle,
            "date_from": self.date_from,
            "date_to": self.date_to,
            "rules": self.rules.to_dict(),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "MetaTopicDefinition":
        return cls(
            id=str(data.get("id", str(uuid.uuid4())[:8])),
            name=str(data.get("name", "Untitled Topic")).strip(),
            description=str(data.get("description", "")).strip(),
            color=str(data.get("color", "#10B981")).strip(),
            target_wgs=[str(w).upper().strip() for w in data.get("target_wgs", ["SA2", "RAN2", "CT1"]) if str(w).strip()],
            cycle_mode=str(data.get("cycle_mode", "past_n_cycles")),
            past_n_cycles=int(data.get("past_n_cycles", 4)),
            from_cycle=str(data.get("from_cycle", "2025-Q1")),
            to_cycle=str(data.get("to_cycle", "2026-Q1")),
            date_from=str(data.get("date_from", "")),
            date_to=str(data.get("date_to", "")),
            rules=MetaTopicRules.from_dict(data.get("rules", {})),
        )

    def matches_tdoc(
        self,
        tdoc_id: str,
        wg: str,
        ai: str = "",
        title: str = "",
        related_wis: str = "",
        abstract: str = "",
        ai_description: str = "",
    ) -> bool:
        """
        Determines if a candidate contribution matches this topic definition.
        Respects explicit manual pins and exclusions first, then evaluates WI tokens,
        WG AI patterns, and keywords (including AI descriptions).
        """
        tid = str(tdoc_id or "").strip().upper()
        if not tid:
            return False

        # 1. Pinned TDocs always match
        if tid in self.rules.pinned_tdocs:
            return True

        # 2. Excluded TDocs never match
        if tid in self.rules.excluded_tdocs:
            return False

        # 3. WG Filter (if target_wgs specified)
        norm_wg = str(wg or "").strip().upper()
        if self.target_wgs and norm_wg:
            if not any(target.upper() == norm_wg for target in self.target_wgs):
                return False

        # 4. Work Item Token intersection
        if self.rules.wi_tokens and related_wis:
            doc_tokens = _extract_base_tokens(related_wis)
            target_tokens = set(self.rules.wi_tokens)
            if doc_tokens.intersection(target_tokens):
                return True

        # 5. WG-specific Agenda Item pattern match
        norm_ai = str(ai or "").strip().lower()
        clean_ai = re.sub(r'^(?:ai|agenda\s*item)\s*[:#-]?\s*', '', norm_ai).strip().rstrip('.')
        if clean_ai and norm_wg in self.rules.ai_patterns:
            patterns = self.rules.ai_patterns[norm_wg]
            for pat in patterns:
                clean_pat = pat.lower().strip()
                if not clean_pat:
                    continue
                clean_pat = re.sub(r'^(?:ai|agenda\s*item)\s*[:#-]?\s*', '', clean_pat).strip().rstrip('.')
                if not clean_pat.endswith("*") and not clean_pat.endswith("?"):
                    # Prefix or exact match: e.g. 20.6.21 matches 20.6.21 or 20.6.21.1
                    if clean_ai == clean_pat or clean_ai.startswith(clean_pat + "."):
                        return True
                elif fnmatch.fnmatch(clean_ai, clean_pat):
                    return True

        # 6. Keyword matching in title, abstract, and agenda item description
        if self.rules.keywords:
            haystack = f"{title} {abstract} {ai_description}".lower()
            for kw in self.rules.keywords:
                clean_kw = kw.strip().lower()
                if clean_kw and clean_kw in haystack:
                    return True

        return False


class MetaTopicManager:
    """Manages loading, saving, and querying persistent Meta-Topics."""

    def __init__(self, config_path: Optional[Path] = None):
        if config_path:
            self.config_path = Path(config_path)
        else:
            root = get_project_root()
            self.config_path = root / DEFAULT_CONFIG_PATH

        self.topics: Dict[str, MetaTopicDefinition] = {}
        self.load()

    def load(self) -> None:
        """Loads topics from meta_topics.json, falling back to default presets if missing."""
        ext_path = to_extended_path(self.config_path)
        if not os.path.exists(ext_path):
            logger.info(f"[MetaTopicManager] Config not found at {self.config_path}, creating defaults.")
            self._create_default_config()
            return

        try:
            with open(ext_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            raw_topics = data.get("topics", [])
            self.topics = {}
            for item in raw_topics:
                topic = MetaTopicDefinition.from_dict(item)
                self.topics[topic.id] = topic
            logger.info(f"[MetaTopicManager] Loaded {len(self.topics)} topic(s) from {self.config_path}")
        except Exception as e:
            logger.error(f"[MetaTopicManager] Failed to load {self.config_path}: {e}")
            self.topics = {}

    def save(self) -> bool:
        """Atomically saves current topics to JSON."""
        ext_path = to_extended_path(self.config_path)
        os.makedirs(os.path.dirname(ext_path), exist_ok=True)
        temp_path = f"{ext_path}.tmp"
        payload = {
            "version": 1,
            "topics": [t.to_dict() for t in self.topics.values()],
        }
        try:
            with open(temp_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2, ensure_ascii=False)
            os.replace(temp_path, ext_path)
            logger.info(f"[MetaTopicManager] Saved {len(self.topics)} topic(s) to {self.config_path}")
            return True
        except Exception as e:
            logger.error(f"[MetaTopicManager] Failed to save {self.config_path}: {e}")
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            return False

    def list_topics(self) -> List[MetaTopicDefinition]:
        return list(self.topics.values())

    def get_topic(self, topic_id: str) -> Optional[MetaTopicDefinition]:
        return self.topics.get(topic_id)

    def save_topic(self, topic: MetaTopicDefinition) -> bool:
        if not topic.id:
            topic.id = re.sub(r'[^a-zA-Z0-9_-]+', '_', topic.name.lower()).strip('_') or str(uuid.uuid4())[:8]
        self.topics[topic.id] = topic
        return self.save()

    def delete_topic(self, topic_id: str) -> bool:
        if topic_id in self.topics:
            del self.topics[topic_id]
            return self.save()
        return False

    def clone_topic(self, topic_id: str) -> Optional[MetaTopicDefinition]:
        orig = self.get_topic(topic_id)
        if not orig:
            return None
        new_topic = copy.deepcopy(orig)
        new_topic.id = f"{orig.id}_copy_{str(uuid.uuid4())[:4]}"
        new_topic.name = f"{orig.name} (Copy)"
        self.topics[new_topic.id] = new_topic
        self.save()
        return new_topic

    def _create_default_config(self) -> None:
        """Initializes default presets."""
        self.topics = {
            "rel19_ambient_iot": MetaTopicDefinition(
                id="rel19_ambient_iot",
                name="Ambient IoT (Rel-19)",
                description="Cross-WG tracking of passive/zero-energy IoT across SA2, RAN2/3, and CT1",
                color="#10B981",
                target_wgs=["SA2", "RAN2", "RAN3", "CT1"],
                cycle_mode="past_n_cycles",
                past_n_cycles=4,
                from_cycle="2025-Q1",
                to_cycle="2026-Q1",
                rules=MetaTopicRules(
                    wi_tokens=["ambientiot", "ambientiot_ph2", "fs_ambientiot"],
                    ai_patterns={
                        "SA2": ["5.14*", "5.12*", "5.8*"],
                        "RAN2": ["8.7*"],
                        "RAN3": ["9.3*"],
                        "CT1": ["6.18*"],
                    },
                    keywords=["ambient iot", "zero power", "passive iot", "energy harvesting", "device discovery"],
                ),
            ),
            "rel19_aiml_ph2": MetaTopicDefinition(
                id="rel19_aiml_ph2",
                name="AI/ML Phase 2 (Rel-19)",
                description="Cross-WG tracking of Artificial Intelligence and Machine Learning enhancements across Core, RAN, and UE",
                color="#6366F1",
                target_wgs=["SA2", "RAN1", "RAN2", "RAN3", "CT1"],
                cycle_mode="past_n_cycles",
                past_n_cycles=4,
                from_cycle="2025-Q1",
                to_cycle="2026-Q1",
                rules=MetaTopicRules(
                    wi_tokens=["aiml", "aiml_ph2", "fs_aiml", "fs_aiml_ph2", "nr_aiml_air_ph2", "nr_aiml_ngran_ph2", "5g_aiml"],
                    ai_patterns={
                        "SA2": ["5.2*", "5.4*"],
                        "RAN1": ["8.1*"],
                        "RAN2": ["8.4*"],
                        "RAN3": ["9.2*"],
                        "CT1": ["6.15*"],
                    },
                    keywords=["ai/ml", "machine learning", "model transfer", "inference", "federated learning", "data collection"],
                ),
            ),
        }
        self.save()
