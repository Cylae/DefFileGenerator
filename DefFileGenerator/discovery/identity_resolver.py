"""
Equipment Identity Resolver for Autonomous Modbus Discovery.
Tokenizes queries, reconciles manufacturer aliases, determines equipment category,
and flags model ambiguities requiring variant clarification.
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

from DefFileGenerator.discovery.models import IdentityCandidate

MANUFACTURER_ALIASES: dict[str, str] = {
    "aiswei": "Solplanet",
    "solplanet": "Solplanet",
    "huawei": "Huawei",
    "sma": "SMA",
    "sma solar": "SMA",
    "sungrow": "Sungrow",
    "goodwe": "GoodWe",
    "foxess": "FoxESS",
    "fox ess": "FoxESS",
    "dyness": "Dyness",
    "cyberpower": "CyberPower",
    "cyber power": "CyberPower",
    "abb": "ABB",
    "solax": "Solax",
    "solax power": "Solax",
    "solis": "Solis",
    "ginlong": "Solis",
    "sofar": "SofarSolar",
    "sofarsolar": "SofarSolar",
    "fronius": "Fronius",
    "kaco": "KACO",
    "schneider": "Schneider",
    "socomec": "Socomec",
    "carlo gavazzi": "CarloGavazzi",
    "carlogavazzi": "CarloGavazzi",
    "alground": "Algodue",
    "algodue": "Algodue",
    "chint": "CHINT",
    "kipp & zonen": "KippZonen",
    "kippzonen": "KippZonen",
    "hukseflux": "Hukseflux",
    "webdyn": "Webdyn",
}

CATEGORY_KEYWORDS: dict[str, str] = {
    "logger": "COM",
    "ailogger": "COM",
    "ai-logger": "COM",
    "ezlogger": "COM",
    "smartlogger": "COM",
    "gateway": "COM",
    "passerelle": "COM",
    "datamanager": "COM",
    "meter": "MET",
    "compteur": "MET",
    "analyser": "MET",
    "analyseur": "MET",
    "energy meter": "MET",
    "power meter": "MET",
    "battery": "BAT",
    "bess": "BAT",
    "batterie": "BAT",
    "storage": "BAT",
    "sensor": "SEN",
    "pyranometer": "SEN",
    "pyranomètre": "SEN",
    "weather": "SEN",
    "temperature": "SEN",
    "sonde": "SEN",
    "inverter": "INV",
    "onduleur": "INV",
    "hybrid": "INV",
    "hybride": "INV",
}


class IdentityResolver:
    """Resolves and tokenizes equipment references into validated identity candidates."""

    def __init__(self, catalog_path: Path | str | None = None) -> None:
        self.catalog_path = (
            catalog_path
            or Path(__file__).resolve().parent.parent.parent / "production" / "CATALOG.json"
        )
        self._known_models_by_mfg: dict[str, list[str]] = {}
        self._load_catalog()

    def _load_catalog(self) -> None:
        if isinstance(self.catalog_path, (str, Path)) and os.path.exists(self.catalog_path):
            try:
                with open(self.catalog_path, encoding="utf-8") as f:
                    data = json.load(f)
                items = data.get("items", []) or data.get("definitions", [])
                for entry in items:
                    mfg = entry.get("manufacturer", "").strip()
                    model = entry.get("model", "").strip()
                    if mfg and model:
                        self._known_models_by_mfg.setdefault(mfg.lower(), []).append(model)
            except Exception:
                pass

    def resolve(self, query: str) -> list[IdentityCandidate]:
        """
        Parses equipment query and returns ranked candidates.
        """
        raw_query = query.strip()
        if not raw_query:
            return []

        clean = re.sub(r"[^\w\s\-\.\/]", " ", raw_query)
        tokens = clean.split()
        if not tokens:
            return []

        # 1. Identify manufacturer
        matched_mfg: str | None = None
        lower_query = raw_query.lower()

        # Check multi-word manufacturer aliases first
        for alias, canon in sorted(MANUFACTURER_ALIASES.items(), key=lambda x: -len(x[0])):
            if lower_query.startswith(alias) or f" {alias} " in f" {lower_query} ":
                matched_mfg = canon
                break

        if not matched_mfg:
            # Fallback to first token
            matched_mfg = tokens[0].capitalize()

        # 2. Extract model string by stripping manufacturer token
        model_part = raw_query
        for alias in MANUFACTURER_ALIASES:
            if MANUFACTURER_ALIASES[alias] == matched_mfg:
                pattern = re.compile(re.escape(alias), re.IGNORECASE)
                model_part = pattern.sub("", model_part)

        model_clean = model_part.strip(" -_:/")
        if not model_clean and len(tokens) > 1:
            model_clean = " ".join(tokens[1:])

        # 3. Check for model ambiguities against catalog
        known_models = self._known_models_by_mfg.get(matched_mfg.lower(), [])
        matching_known = [
            m
            for m in known_models
            if model_clean.lower() in m.lower() or m.lower() in model_clean.lower()
        ]

        # Distinguish hardware revisions or variant suffixes (e.g., G2, Pro, Plus, HV)
        is_ambiguous = False
        distinguishing_factors: list[str] = []
        candidates: list[IdentityCandidate] = []

        if len(matching_known) > 1:
            # Check if matching models represent different series
            distinct_series = {re.sub(r"[\d\-_]", "", m) for m in matching_known}
            if len(distinct_series) > 1 or len(matching_known) >= 2:
                is_ambiguous = True
                distinguishing_factors = matching_known[:5]

        # Primary candidate
        candidates.append(
            IdentityCandidate(
                manufacturer=matched_mfg,
                model_family=tokens[1] if len(tokens) > 1 else model_clean,
                exact_model=model_clean or "Generic",
                confidence_score=0.95 if not is_ambiguous else 0.70,
                is_ambiguous=is_ambiguous,
                distinguishing_factors=distinguishing_factors,
            )
        )

        return candidates
