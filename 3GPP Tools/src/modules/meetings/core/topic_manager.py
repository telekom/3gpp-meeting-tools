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


def get_default_meta_topics_path() -> Path:
    """Path to the built-in factory meta-topics (`src/config/meta_topics.json`, tracked in Git)."""
    return get_project_root() / DEFAULT_CONFIG_PATH


def get_user_meta_topics_path() -> Path:
    """Path to local user meta-topics configuration (`src/config/meta_topics_user.json`, Git-ignored)."""
    return get_project_root() / USER_CONFIG_PATH


def get_app_data_meta_topics_path() -> Path:
    """Path to persistent machine-level user meta-topics (`~/3GPP_Delegate_Helper/meta_topics_user.json`)."""
    from core.utils.paths import get_app_data_root
    return get_app_data_root() / "meta_topics_user.json"


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


def _get_builtin_presets() -> Dict[str, MetaTopicDefinition]:
    """Fallback canonical 3GPP presets defined directly in Python."""
    return {
        "rel19_ambient_iot": MetaTopicDefinition(
            id="rel19_ambient_iot",
            name="Ambient IoT (Rel-19)",
            description="Cross-WG tracking of passive/zero-energy IoT across SA2 (architecture), RAN2/3 (radio & protocol), and CT1 (NAS/URSP)",
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
        "rel20_6g_architecture": MetaTopicDefinition(
            id="rel20_6g_architecture",
            name="6G Architecture & Feasibility (Rel-20)",
            description="Studies on 6G radio and system architecture, requirements, and foundational principles",
            color="#F59E0B",
            target_wgs=["SA2", "RAN2", "RAN3", "CT1"],
            cycle_mode="past_n_cycles",
            past_n_cycles=4,
            from_cycle="2025-Q1",
            to_cycle="2026-Q1",
            rules=MetaTopicRules(
                wi_tokens=["fs_6g_arc", "6g_arc", "fs_6g_radio", "6gr", "6g"],
                ai_patterns={
                    "SA2": ["5.1*", "5.2*"],
                    "RAN2": ["8.1*"],
                    "RAN3": ["10*"],
                    "CT1": ["6.1*"],
                },
                keywords=["6g architecture", "study on 6g", "sub-terahertz", "integrated sensing", "isac", "compute network"],
            ),
        ),
        "rel19_network_energy_savings": MetaTopicDefinition(
            id="rel19_network_energy_savings",
            name="Network Energy Savings (Rel-19)",
            description="Network energy savings across radio access, gNB sleep states, and core network coordination",
            color="#06B6D4",
            target_wgs=["SA2", "RAN1", "RAN2", "RAN3"],
            cycle_mode="past_n_cycles",
            past_n_cycles=4,
            from_cycle="2025-Q1",
            to_cycle="2026-Q1",
            rules=MetaTopicRules(
                wi_tokens=["nes", "nes_ph2", "fs_nes", "fs_nes_ph2", "nr_nes_ph2"],
                ai_patterns={
                    "SA2": ["5.9*"],
                    "RAN1": ["8.3*"],
                    "RAN2": ["8.6*"],
                    "RAN3": ["9.4*"],
                },
                keywords=["energy savings", "nes", "power consumption", "cell sleep", "rf shutdown"],
            ),
        ),
        "rel20_data_framework": MetaTopicDefinition(
            id="rel20_data_framework",
            name="Data Framework & Data Transfer (Rel-20)",
            description="6G E2E Data Framework and Data Transfer topic across SA2, RAN2, RAN3, SA5, and CT4",
            color="#0284C7",
            target_wgs=["SA2", "RAN2", "RAN3", "SA5", "CT4"],
            cycle_mode="past_n_cycles",
            past_n_cycles=4,
            from_cycle="2025-Q1",
            to_cycle="2026-Q4",
            rules=MetaTopicRules(
                wi_tokens=["6g_df", "fs_6g_df", "fs_6g_dfcn_ct", "6g_dfcn_ct", "fs_6g_dfcn"],
                ai_patterns={
                    "SA2": ["20.6.21*"],
                    "RAN2": ["9.3.3.1*", "10.3.3.1*"],
                    "RAN3": ["10.2.3*", "10.2.1*"],
                    "SA5": ["6.20.6.3*", "5.3*"],
                    "CT4": ["21.5*"],
                },
                keywords=["data framework", "data transfer", "ki#21", "key issue #21", "key issue 21", "6g data framework", "data management framework"],
            ),
        ),
    }


class MetaTopicManager:
    """Manages loading, saving, and querying persistent Meta-Topics."""

    def __init__(self, config_path: Optional[Path] = None):
        self.custom_config_path: Optional[Path] = Path(config_path) if config_path else None
        if self.custom_config_path:
            self.config_path = self.custom_config_path
            self.user_config_path = self.custom_config_path
            self.app_data_config_path = None
        else:
            self.config_path = get_default_meta_topics_path()
            self.user_config_path = get_user_meta_topics_path()
            self.app_data_config_path = get_app_data_meta_topics_path()

        self.topics: Dict[str, MetaTopicDefinition] = {}
        self.load()

    def _load_from_path(self, path: Path) -> Tuple[bool, int]:
        """Loads topics dictionary from a specific JSON file."""
        ext_path = to_extended_path(path)
        if not os.path.exists(ext_path):
            return False, 0
        try:
            with open(ext_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            raw_topics = data.get("topics", [])
            loaded: Dict[str, MetaTopicDefinition] = {}
            for item in raw_topics:
                topic = MetaTopicDefinition.from_dict(item)
                loaded[topic.id] = topic
            self.topics = loaded
            return True, len(loaded)
        except Exception as e:
            logger.error(f"[MetaTopicManager] Failed to load {path}: {e}")
            return False, 0

    def load(self) -> None:
        """
        Loads topics respecting user persistence:
          1. Custom config path (if explicitly provided, e.g. for unit tests)
          2. Workspace user topics file (`src/config/meta_topics_user.json`, Git-ignored)
          3. AppData user topics file (`~/3GPP_Delegate_Helper/meta_topics_user.json`)
          4. Default factory presets (`src/config/meta_topics.json`), copied to user config on first run
        """
        if self.custom_config_path:
            if not self.custom_config_path.exists():
                self.topics = _get_builtin_presets()
                self.save()
            else:
                self._load_from_path(self.custom_config_path)
            return

        # 1. Check workspace user file
        if self.user_config_path.exists():
            ok, count = self._load_from_path(self.user_config_path)
            if ok:
                logger.info(f"[MetaTopicManager] Loaded {count} topic(s) from user config {self.user_config_path}")
                if self.app_data_config_path and not self.app_data_config_path.exists():
                    self._save_to_path(self.app_data_config_path)
                return

        # 2. Check AppData user file
        if self.app_data_config_path and self.app_data_config_path.exists():
            ok, count = self._load_from_path(self.app_data_config_path)
            if ok:
                logger.info(f"[MetaTopicManager] Loaded {count} topic(s) from AppData {self.app_data_config_path}")
                self._save_to_path(self.user_config_path)
                return

        # 3. First run ever: neither user config exists. Seed from factory presets or defaults.
        logger.info("[MetaTopicManager] No user config found; initializing from factory presets.")
        if self.config_path.exists():
            ok, count = self._load_from_path(self.config_path)
            if not ok or count == 0:
                self.topics = _get_builtin_presets()
        else:
            self.topics = _get_builtin_presets()

        self.save()

    def _save_to_path(self, path: Path) -> bool:
        """Atomically saves current topics to a specific JSON file."""
        ext_path = to_extended_path(path)
        os.makedirs(os.path.dirname(ext_path), exist_ok=True)
        temp_path = f"{ext_path}.tmp"
        payload = {
            "version": 1,
            "topics": [t.to_dict() for t in self.topics.values()],
        }
        try:
            with open(temp_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2, ensure_ascii=False)

            # Windows retry loop guarding against transient OneDrive / antivirus file locks
            max_attempts = 15
            for attempt in range(max_attempts):
                try:
                    os.replace(temp_path, ext_path)
                    break
                except (PermissionError, OSError):
                    if attempt == max_attempts - 1:
                        import shutil
                        shutil.copyfile(temp_path, ext_path)
                        try:
                            os.remove(temp_path)
                        except Exception:
                            pass
                        break
                    import time
                    time.sleep(0.1)

            logger.info(f"[MetaTopicManager] Saved {len(self.topics)} topic(s) to {path}")
            return True
        except Exception as e:
            logger.error(f"[MetaTopicManager] Failed to save {path}: {e}")
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            return False

    def save(self) -> bool:
        """Atomically saves current topics to the user configuration file(s)."""
        target_path = self.custom_config_path if self.custom_config_path else self.user_config_path
        ok = self._save_to_path(target_path)
        if ok and not self.custom_config_path and self.app_data_config_path:
            self._save_to_path(self.app_data_config_path)
        return ok

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
            logger.info(f"[MetaTopicManager] Deleted topic '{topic_id}' (remaining: {len(self.topics)})")
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

    def restore_default_topics(self, overwrite: bool = False) -> int:
        """
        Restores built-in default topics from meta_topics.json or code presets.
        If overwrite is False, only adds built-in topics whose IDs are not currently present.
        If overwrite is True, replaces self.topics entirely with the built-in topics.
        Returns the number of topics restored.
        """
        factory_topics = self._read_factory_presets()
        if overwrite:
            self.topics = copy.deepcopy(factory_topics)
            count = len(factory_topics)
        else:
            count = 0
            for tid, t in factory_topics.items():
                if tid not in self.topics:
                    self.topics[tid] = copy.deepcopy(t)
                    count += 1
        if count > 0 or overwrite:
            self.save()
        return count

    def _read_factory_presets(self) -> Dict[str, MetaTopicDefinition]:
        """Reads factory presets from config/meta_topics.json or Python code fallback."""
        factory_path = get_default_meta_topics_path()
        if factory_path.exists():
            ext_path = to_extended_path(factory_path)
            try:
                with open(ext_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                raw = data.get("topics", [])
                result = {}
                for item in raw:
                    t = MetaTopicDefinition.from_dict(item)
                    result[t.id] = t
                if result:
                    return result
            except Exception as e:
                logger.error(f"[MetaTopicManager] Failed to read factory presets from {factory_path}: {e}")
        return _get_builtin_presets()
