"""
SunSpec Relevance Evaluator and Intelligent Architectural Slicer.

Determines conditionally before generation whether one or more SunSpec definition
files should be produced, enforcing strict gating rules:
  - Generates iff explicit, actionable SunSpec Model IDs, PICS, or SunS markers exist.
  - Skips cleanly if source is purely proprietary Modbus (e.g. Huawei 30000..39999).
  - Skips cleanly if 'SunSpec' is a commercial/marketing mention without models or PICS.
  - Skips cleanly if all points in detected blocks are marked unimplemented.
  - Partitions incompatible or optional profiles into distinct definition targets:
      1. Legacy Profile (1 + 101/102/103 + 120-126 + 160) vs DER 700 Profile (1 + 701 + 702 + 704 + 714).
         Never mixes 103 and 701 in the same definition file.
      2. AC Topologies: Single-Phase (101/201), Split-Phase (102/202), Three-Phase (103/203/204).
      3. Functional slices: Core AC/DC vs Storage (124/802) vs Regulation/Controls (123/704).
      4. Deterministic Webdyn naming complying with production governor.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from DefFileGenerator.sunspec.constants import (
    BASE_ADDRESS_0,
    BASE_ADDRESS_1,
    COMMON_MODEL_LEN_STANDARD,
    DER_700_MODELS,
    INVERTER_CONTROL_MODELS,
    LEGACY_INVERTER_MODELS,
    LEGACY_INVERTER_SINGLE_PHASE,
    LEGACY_INVERTER_SPLIT_PHASE,
    LEGACY_INVERTER_THREE_PHASE,
    METER_MODELS,
    METER_SINGLE_PHASE,
    METER_SPLIT_PHASE,
    METER_THREE_PHASE_DELTA,
    METER_THREE_PHASE_WYE,
    MPPT_STRING_MODELS,
    STORAGE_BESS_MODELS,
    SUNS_MAGIC_HEX,
    SUNS_MAGIC_STR,
    SunSpecCliMode,
    SunSpecProfileGeneration,
    SunSpecTopology,
)
from DefFileGenerator.sunspec.pics_parser import PICSDocument, PICSParser

logger = logging.getLogger(__name__)

RE_CLEAN_NAME = re.compile(r"[^A-Za-z0-9]+")


@dataclass
class SunSpecGenerationTarget:
    """Definition of a single, coherent SunSpec definition file to generate."""
    prefix: str                          # "INV", "MET", "BAT"
    manufacturer: str
    model_name: str
    profile: SunSpecProfileGeneration
    topology: SunSpecTopology
    model_ids: list[int]
    base_address: int = BASE_ADDRESS_0
    common_model_length: int = COMMON_MODEL_LEN_STANDARD
    repeating_counts: dict[int, int] = field(default_factory=dict)
    functional_scope: str = "Core"      # "Core", "Storage", "Control"
    output_filename: str = ""

    def get_default_filename(self) -> str:
        """Generates a deterministic filename satisfying Apollo 5 Pillars naming rules."""
        mfg = RE_CLEAN_NAME.sub("-", self.manufacturer).strip("-") or "Generic"
        mdl = RE_CLEAN_NAME.sub("-", self.model_name).strip("-") or "SunSpec"

        # Construct specific profile and topology descriptors
        parts = [self.prefix, mfg, mdl, "SunSpec", self.profile.value]
        if self.topology not in (SunSpecTopology.UNKNOWN, SunSpecTopology.STORAGE, SunSpecTopology.CONTROL):
            parts.append(self.topology.value)
        elif self.functional_scope and self.functional_scope != "Core":
            parts.append(self.functional_scope)

        return "_".join(parts) + ".csv"


@dataclass
class SunSpecRelevanceDecision:
    """Outcome of the SunSpec relevance evaluation gate."""
    should_generate: bool
    reason: str
    detected_profiles: list[SunSpecProfileGeneration] = field(default_factory=list)
    detected_models: list[int] = field(default_factory=list)
    targets: list[SunSpecGenerationTarget] = field(default_factory=list)
    is_forced: bool = False
    is_skipped: bool = False
    pics_doc: PICSDocument | None = None


class SunSpecRelevanceEvaluator:
    """Evaluates input documentation and plans conditional SunSpec definition generation."""

    def __init__(self) -> None:
        self.parser = PICSParser()

    def evaluate(
        self,
        source_path: Path | str | None = None,
        raw_text: str | None = None,
        intermediate_rows: list[dict[str, Any]] | None = None,
        pics_doc: PICSDocument | None = None,
        cli_mode: str | SunSpecCliMode = SunSpecCliMode.AUTO,
        manufacturer: str | None = None,
        model: str | None = None,
        category: str | None = None,
        forced_base_address: int | None = None,
        forced_ports: int | None = None,
    ) -> SunSpecRelevanceDecision:
        """
        Evaluates whether SunSpec definition files should be generated.

        Args:
            source_path: Path to source document (Excel PICS, PDF, CSV, JSON).
            raw_text: Raw textual content extracted from documentation.
            intermediate_rows: Intermediate extracted register rows if already available.
            pics_doc: Pre-parsed PICSDocument if available.
            cli_mode: CLI control mode ("auto", "force", "off").
            manufacturer: Manufacturer override.
            model: Equipment model override.
            category: Equipment category override ("Inverter", "Meter", "Battery").
            forced_base_address: Force Base 40000 or 40001.
            forced_ports: Force number of MPPT strings or DC ports.

        Returns:
            SunSpecRelevanceDecision detailing the gate outcome and planned targets.
        """
        mode = SunSpecCliMode(cli_mode.lower()) if isinstance(cli_mode, str) else cli_mode

        # Gate 0: Explicit CLI OFF
        if mode == SunSpecCliMode.OFF:
            logger.info("SunSpec evaluation skipped: explicitly disabled via --sunspec=off.")
            return SunSpecRelevanceDecision(
                should_generate=False,
                is_skipped=True,
                reason="SunSpec generation explicitly disabled via --sunspec=off",
            )

        # 1. Parse PICS if not provided
        if pics_doc is None and source_path:
            p_src = Path(source_path)
            if p_src.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm", ".csv", ".json"):
                pics_doc = self.parser.parse_file(p_src)

        # Gate 1: Explicit CLI FORCE
        if mode == SunSpecCliMode.FORCE:
            return self._plan_forced_generation(
                source_path=source_path,
                pics_doc=pics_doc,
                manufacturer=manufacturer,
                model=model,
                category=category,
                forced_base_address=forced_base_address,
                forced_ports=forced_ports,
            )

        # 2. Gate: Marketing-only SunSpec mention check
        if self._is_marketing_only_mention(raw_text, pics_doc, intermediate_rows):
            logger.info(
                "SunSpec evaluation skipped: keyword 'SunSpec' appears only as generic/marketing "
                "mention without actionable Model IDs or PICS."
            )
            return SunSpecRelevanceDecision(
                should_generate=False,
                is_skipped=True,
                reason="SunSpec keyword appears only in generic/commercial marketing context without actionable models or PICS",
                pics_doc=pics_doc,
            )

        # 3. Gate: Purely proprietary Modbus check
        if self._is_pure_proprietary_modbus(source_path, raw_text, intermediate_rows, pics_doc):
            logger.info(
                "SunSpec evaluation skipped: source describes only proprietary Modbus map; "
                "no SunSpec Model IDs, PICS, or SunS marker found."
            )
            return SunSpecRelevanceDecision(
                should_generate=False,
                is_skipped=True,
                reason="Source describes only proprietary Modbus map; no SunSpec Model IDs or PICS identified",
                pics_doc=pics_doc,
            )

        # 4. Actionable SunSpec models inspection
        if pics_doc is not None and pics_doc.has_models:
            operational_models = [m for m in pics_doc.models.values() if m.model_id > 1]
            if operational_models and all(not m.is_supported for m in operational_models):
                logger.info("SunSpec evaluation skipped: all operational points in detected models are marked unimplemented.")
                return SunSpecRelevanceDecision(
                    should_generate=False,
                    is_skipped=True,
                    reason="All points in detected SunSpec models are marked unimplemented or reserved in PICS",
                    pics_doc=pics_doc,
                )

            # Plan intelligent targets based on verified PICS models
            return self._plan_targets_from_pics(
                pics_doc=pics_doc,
                manufacturer_override=manufacturer,
                model_override=model,
                category_override=category,
                forced_base_address=forced_base_address,
                forced_ports=forced_ports,
            )

        # 5. Check intermediate rows for SunS marker or Model ID headers
        if intermediate_rows:
            models_from_rows = self._detect_models_from_rows(intermediate_rows)
            if models_from_rows:
                return self._plan_targets_from_models(
                    models=models_from_rows,
                    source_path=source_path,
                    manufacturer=manufacturer,
                    model=model,
                    category=category,
                    forced_base_address=forced_base_address,
                    forced_ports=forced_ports,
                )

        # 6. Check raw text for explicit Model IDs
        if raw_text:
            models_from_text = self._detect_models_from_text(raw_text)
            if models_from_text:
                return self._plan_targets_from_models(
                    models=models_from_text,
                    source_path=source_path,
                    manufacturer=manufacturer,
                    model=model,
                    category=category,
                    forced_base_address=forced_base_address,
                    forced_ports=forced_ports,
                )

        # Default fallback: No actionable SunSpec detected
        return SunSpecRelevanceDecision(
            should_generate=False,
            is_skipped=True,
            reason="No actionable SunSpec Model IDs, PICS, or SunS marker detected in source",
            pics_doc=pics_doc,
        )

    def _is_marketing_only_mention(
        self,
        raw_text: str | None,
        pics_doc: PICSDocument | None,
        intermediate_rows: list[dict[str, Any]] | None,
    ) -> bool:
        """Detects if 'SunSpec' is only a generic brand/marketing mention without actionable models."""
        if not raw_text:
            return False

        has_sunspec_kw = bool(re.search(r"\bsunspec\b", raw_text, re.IGNORECASE))
        if not has_sunspec_kw:
            return False

        # If PICS document has identified models, it is not marketing-only
        if pics_doc and pics_doc.has_models:
            return False

        # If rows contain Model IDs or SunS marker, it is not marketing-only
        if intermediate_rows and self._detect_models_from_rows(intermediate_rows):
            return False

        # If text contains specific SunSpec model IDs or PICS keywords
        has_specific_model = bool(
            re.search(r"\b(?:model|bloc|profile)\s*(?:101|102|103|120|123|124|160|201|202|203|204|701|702|704|714|802)\b", raw_text, re.IGNORECASE)
            or re.search(r"\b4000[0-4]\b", raw_text)
            or re.search(r"\b0x53756E53\b", raw_text, re.IGNORECASE)
        )
        return not has_specific_model

    def _is_pure_proprietary_modbus(
        self,
        source_path: Path | str | None,
        raw_text: str | None,
        intermediate_rows: list[dict[str, Any]] | None,
        pics_doc: PICSDocument | None,
    ) -> bool:
        """Detects if the source is purely proprietary Modbus without SunSpec points."""
        if pics_doc and pics_doc.has_models:
            return False

        # If intermediate rows exist, check register addresses
        if intermediate_rows:
            addrs: list[int] = []
            for r in intermediate_rows:
                a_val = r.get("Address") or r.get("address")
                try:
                    if a_val is not None:
                        addrs.append(int(str(a_val).split("_")[0]))
                except ValueError:
                    pass

            if addrs:
                # Proprietary maps like Huawei (30000..39999), Sungrow (4999..13000), etc.
                has_sunspec_base = any(40000 <= a <= 40500 for a in addrs)
                if not has_sunspec_base:
                    return True

        if raw_text:
            # Check for pure proprietary headers
            if re.search(r"\b(huawei|sungrow|kaco|chint|victron|socomec)\b", raw_text, re.IGNORECASE):
                if not re.search(r"\bsunspec\b", raw_text, re.IGNORECASE) and not re.search(r"\b4000[0-2]\b", raw_text):
                    return True

        return False

    def _detect_models_from_rows(self, rows: list[dict[str, Any]]) -> list[int]:
        """Scans extracted register rows for SunSpec Model IDs."""
        detected = set()
        for r in rows:
            name = str(r.get("Name") or r.get("name") or "").strip()
            desc = str(r.get("Description") or r.get("description") or "").strip()
            addr_str = str(r.get("Address") or r.get("address") or "")
            m = re.search(r"\b(?:sunspec\s+)?model\s*(\d{1,4})\b", f"{name} {desc}", re.IGNORECASE)
            if m:
                detected.add(int(m.group(1)))

            if addr_str in ("40000", "40001") and ("suns" in name.lower() or "suns" in desc.lower()):
                detected.add(1)

        return sorted(detected)

    def _detect_models_from_text(self, text: str) -> list[int]:
        """Scans raw text for explicit SunSpec Model IDs."""
        if not text:
            return []
        detected = set()
        matches = re.finditer(r"\b(?:sunspec\s+)?(?:model|bloc|profile)\s*(\d{1,4})\b", text, re.IGNORECASE)
        for m in matches:
            mid = int(m.group(1))
            if mid in (1, 101, 102, 103, 120, 123, 124, 160, 201, 202, 203, 204, 701, 702, 704, 714, 802):
                detected.add(mid)
        if re.search(r"\b0x53756E53\b", text, re.IGNORECASE) or re.search(r"\b4000[0-2]\b", text):
            detected.add(1)
        return sorted(detected)

    def _plan_targets_from_pics(
        self,
        pics_doc: PICSDocument,
        manufacturer_override: str | None,
        model_override: str | None,
        category_override: str | None,
        forced_base_address: int | None,
        forced_ports: int | None,
    ) -> SunSpecRelevanceDecision:
        """Constructs split definition targets from a verified PICSDocument."""
        mfg = manufacturer_override or pics_doc.manufacturer or "Generic"
        mdl = model_override or pics_doc.model_name or "Device"
        base_addr = forced_base_address if forced_base_address is not None else pics_doc.base_address
        l_common = pics_doc.common_model_length

        available_models = set(pics_doc.models.keys())
        targets: list[SunSpecGenerationTarget] = []
        detected_profiles: list[SunSpecProfileGeneration] = []

        # Repeating counts override if provided
        repeating_counts = dict(pics_doc.repeating_counts)
        if forced_ports is not None:
            if 160 in available_models:
                repeating_counts[160] = forced_ports
            if 714 in available_models:
                repeating_counts[714] = forced_ports

        # ----------------------------------------------------------------------
        # Target 1: Legacy Inverter Profile (Models 1 + 101/102/103 + opt 120-126 + 160)
        # ----------------------------------------------------------------------
        legacy_inverters = available_models.intersection(LEGACY_INVERTER_MODELS)
        if legacy_inverters:
            detected_profiles.append(SunSpecProfileGeneration.LEGACY)
            # Separate by AC topology if multiple are implemented
            for inv_mid in sorted(legacy_inverters):
                topology = (
                    SunSpecTopology.SINGLE_PHASE if inv_mid == 101
                    else (SunSpecTopology.SPLIT_PHASE if inv_mid == 102 else SunSpecTopology.THREE_PHASE)
                )

                core_models = [1, inv_mid]
                if 120 in available_models:
                    core_models.append(120)
                if 123 in available_models:
                    core_models.append(123)
                if 160 in available_models:
                    core_models.append(160)

                target = SunSpecGenerationTarget(
                    prefix="INV",
                    manufacturer=mfg,
                    model_name=mdl,
                    profile=SunSpecProfileGeneration.LEGACY,
                    topology=topology,
                    model_ids=core_models,
                    base_address=base_addr,
                    common_model_length=l_common,
                    repeating_counts=repeating_counts,
                    functional_scope="Core",
                )
                target.output_filename = target.get_default_filename()
                targets.append(target)

            # Separate Storage file if hybrid storage model 124 is present
            if 124 in available_models:
                storage_target = SunSpecGenerationTarget(
                    prefix="BAT",
                    manufacturer=mfg,
                    model_name=mdl,
                    profile=SunSpecProfileGeneration.LEGACY,
                    topology=SunSpecTopology.STORAGE,
                    model_ids=[1, 124],
                    base_address=base_addr,
                    common_model_length=l_common,
                    functional_scope="Storage",
                )
                storage_target.output_filename = storage_target.get_default_filename()
                targets.append(storage_target)

        # ----------------------------------------------------------------------
        # Target 2: DER 700 Series Profile (Models 1 + 701 + opt 702, 704, 714)
        # NEVER mix 701 with 103!
        # ----------------------------------------------------------------------
        der700_models = available_models.intersection(DER_700_MODELS)
        if der700_models and (701 in der700_models or 702 in der700_models):
            detected_profiles.append(SunSpecProfileGeneration.DER700)
            core_der_models = [1]
            if 701 in der700_models:
                core_der_models.append(701)
            if 702 in der700_models:
                core_der_models.append(702)
            if 704 in der700_models:
                core_der_models.append(704)
            if 714 in der700_models:
                core_der_models.append(714)

            target_der = SunSpecGenerationTarget(
                prefix="INV",
                manufacturer=mfg,
                model_name=mdl,
                profile=SunSpecProfileGeneration.DER700,
                topology=SunSpecTopology.THREE_PHASE,
                model_ids=core_der_models,
                base_address=base_addr,
                common_model_length=l_common,
                repeating_counts=repeating_counts,
                functional_scope="DER_Measurement",
            )
            target_der.output_filename = target_der.get_default_filename()
            targets.append(target_der)

        # ----------------------------------------------------------------------
        # Target 3: AC Meters Profile (Models 201..204)
        # ----------------------------------------------------------------------
        meter_models = available_models.intersection(METER_MODELS)
        if meter_models:
            for m_mid in sorted(meter_models):
                topo = (
                    SunSpecTopology.SINGLE_PHASE if m_mid in METER_SINGLE_PHASE
                    else (SunSpecTopology.SPLIT_PHASE if m_mid in METER_SPLIT_PHASE
                          else (SunSpecTopology.THREE_PHASE_DELTA if m_mid in METER_THREE_PHASE_DELTA
                                else SunSpecTopology.THREE_PHASE))
                )
                target_meter = SunSpecGenerationTarget(
                    prefix="MET",
                    manufacturer=mfg,
                    model_name=mdl,
                    profile=SunSpecProfileGeneration.LEGACY,
                    topology=topo,
                    model_ids=[1, m_mid],
                    base_address=base_addr,
                    common_model_length=l_common,
                    functional_scope="Meter",
                )
                target_meter.output_filename = target_meter.get_default_filename()
                targets.append(target_meter)

        # ----------------------------------------------------------------------
        # Target 4: Battery BESS Storage Model 802
        # ----------------------------------------------------------------------
        if 802 in available_models:
            target_bat = SunSpecGenerationTarget(
                prefix="BAT",
                manufacturer=mfg,
                model_name=mdl,
                profile=SunSpecProfileGeneration.LEGACY,
                topology=SunSpecTopology.STORAGE,
                model_ids=[1, 802],
                base_address=base_addr,
                common_model_length=l_common,
                functional_scope="BESS_Storage",
            )
            target_bat.output_filename = target_bat.get_default_filename()
            targets.append(target_bat)

        should_gen = bool(targets)
        reason = (
            f"Confirmed actionable SunSpec models {sorted(available_models)}; "
            f"planned {len(targets)} decoupled definition file(s)"
            if should_gen
            else "Detected SunSpec sheets but none matched active generator profiles"
        )

        return SunSpecRelevanceDecision(
            should_generate=should_gen,
            reason=reason,
            detected_profiles=detected_profiles,
            detected_models=sorted(available_models),
            targets=targets,
            pics_doc=pics_doc,
        )

    def _plan_targets_from_models(
        self,
        models: list[int],
        source_path: Path | str | None,
        manufacturer: str | None,
        model: str | None,
        category: str | None,
        forced_base_address: int | None,
        forced_ports: int | None,
    ) -> SunSpecRelevanceDecision:
        """Constructs targets when models are identified from register rows."""
        # Convert identified list of models to a pseudo-PICSDocument
        doc = PICSDocument(
            file_path=Path(source_path) if source_path else None,
            manufacturer=manufacturer or "Generic",
            model_name=model or "Device",
            category=category or "Inverter",
            base_address=forced_base_address or BASE_ADDRESS_0,
        )
        for mid in models:
            from DefFileGenerator.sunspec.models_registry import get_sunspec_model
            mdef = get_sunspec_model(mid)
            if mdef:
                from DefFileGenerator.sunspec.pics_parser import PICSModel
                doc.models[mid] = PICSModel(
                    model_id=mid,
                    address=40002,
                    length=mdef.standard_length,
                )

        return self._plan_targets_from_pics(
            pics_doc=doc,
            manufacturer_override=manufacturer,
            model_override=model,
            category_override=category,
            forced_base_address=forced_base_address,
            forced_ports=forced_ports,
        )

    def _plan_forced_generation(
        self,
        source_path: Path | str | None,
        pics_doc: PICSDocument | None,
        manufacturer: str | None,
        model: str | None,
        category: str | None,
        forced_base_address: int | None,
        forced_ports: int | None,
    ) -> SunSpecRelevanceDecision:
        """Constructs targets when generation is forced via --sunspec=force."""
        mfg = manufacturer or (pics_doc.manufacturer if pics_doc else None) or "Generic"
        mdl = model or (pics_doc.model_name if pics_doc else None) or "SunSpec-Device"
        cat = category or (pics_doc.category if pics_doc else None) or "Inverter"
        base_addr = forced_base_address or (pics_doc.base_address if pics_doc else None) or BASE_ADDRESS_0
        l_common = pics_doc.common_model_length if pics_doc else COMMON_MODEL_LEN_STANDARD

        if pics_doc and pics_doc.has_models:
            decision = self._plan_targets_from_pics(
                pics_doc=pics_doc,
                manufacturer_override=mfg,
                model_override=mdl,
                category_override=cat,
                forced_base_address=base_addr,
                forced_ports=forced_ports,
            )
            decision.is_forced = True
            decision.reason = f"Forced by CLI flag --sunspec=force (using detected models {decision.detected_models})"
            return decision

        # Default fallback models based on equipment category
        rep_counts = {}
        if cat.lower() == "meter":
            prefix = "MET"
            profile = SunSpecProfileGeneration.LEGACY
            topo = SunSpecTopology.THREE_PHASE
            model_ids = [1, 203]
        elif cat.lower() in ("battery", "storage"):
            prefix = "BAT"
            profile = SunSpecProfileGeneration.LEGACY
            topo = SunSpecTopology.STORAGE
            model_ids = [1, 802]
        else:
            prefix = "INV"
            profile = SunSpecProfileGeneration.LEGACY
            topo = SunSpecTopology.THREE_PHASE
            model_ids = [1, 103, 120, 160]
            rep_counts[160] = forced_ports if forced_ports is not None else 2

        target = SunSpecGenerationTarget(
            prefix=prefix,
            manufacturer=mfg,
            model_name=mdl,
            profile=profile,
            topology=topo,
            model_ids=model_ids,
            base_address=base_addr,
            common_model_length=l_common,
            repeating_counts=rep_counts,
            functional_scope="Core",
        )
        target.output_filename = target.get_default_filename()

        return SunSpecRelevanceDecision(
            should_generate=True,
            reason="Forced by CLI flag --sunspec=force (standard default model template)",
            detected_profiles=[profile],
            detected_models=model_ids,
            targets=[target],
            is_forced=True,
            pics_doc=pics_doc,
        )
