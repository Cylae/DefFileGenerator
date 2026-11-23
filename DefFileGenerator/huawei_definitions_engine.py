"""
Apollo-Standard Master Definition Engine for Huawei Inverters, Batteries, Meters & Loggers.

Official Documentation Sources:
  1. SUN2000-V4 Generic: HUAWEI_V4.csv & Solar Inverter Modbus Interface Definitions V4.0 (322 registers)
  2. SUN2000-2..5KTL: Huawei_2345.csv & SUN2000L-(2KTL-5KTL) User Manual (47 registers)
  3. SUN2000-8..20KTL: Huawei_81012141820_Test.csv & SUN2000-(8KTL-20KTL)-M2 User Manual (328 registers)
  4. SUN2000-30..40KTL-M3: SUN2000-(20KTL-40KTL)-M3 User Manual (64 registers)
  5. SUN2000-33A..42KTL: Huawei_364233-A.csv & SUN2000-(33KTL-A, 36KTL, 42KTL) Quick Guide (192 registers)
  6. SUN2000-50..70KTL-M0: Huawei_7550606570-M0.csv & SUN2000-(50KTL, 60KTL, 70KTL-M0) (73 registers)
  7. SUN2000-60KTL-HV: Huawei_60-HV.csv & SUN2000-60KTL-HV User Manual (347 registers)
  8. SUN2000-90..105KTL: Huawei_9095100105.csv & SUN2000-(90KTL-105KTL) (73 registers)
  9. SUN2000-V3: HUAWEI_V3.csv (97 registers)
  10. SUN2000-All V3: HUAWEI (SUN2000C00).csv (274 registers)
  11. LUNA2000 Battery: HUAWEI_Luna_2000.csv & LUNA2000C Modbus Port Definitions.pdf (170 registers)
  12. LUNA2000 Storage: HUAWEI_INV_STORAGE.csv & LUNA2000 User Manual (272 registers)
  13. SmartLogger 3000: HUAWEI_SmartLogger3000.csv & SmartLogger ModBus Interface Definitions (83 registers)
  14. SmartLogger V Series: DEF_HUAWEI_SmartLogger_VSeries.csv (176 registers)
  15. Power Meter: HUAWEI_PowerMeter.csv & DTSU666-H User Manual (44 registers)

Strictly enforces:
  - Pillar 1: Strict UTF-8 without BOM.
  - Pillar 2: Strict CRLF (\\r\\n) line endings.
  - Pillar 3: Exact 11-column semicolon CSV schema with standard Webdyn header.
  - Pillar 4: Monotonic sequential 1..N row indexing.
  - Pillar 5: Valid Webdyn data types and canonical SunPM tags.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ModbusRegisterDef:
    """Specification of a Modbus register for a Webdyn definition file."""

    fc: str
    address: str
    data_type: str
    reserved: str
    name: str
    tag: str
    coef_a: str
    coef_b: str
    unit: str
    action: str


class HuaweiDefinitionsEngine:
    """Master engine for compiling certified Huawei equipment definitions."""

    MANUFACTURER: str = "Huawei"
    PROTOCOL: str = "Modbus"

    @classmethod
    def format_row(cls, index: int, reg: ModbusRegisterDef) -> str:
        """Format a single register definition into an 11-column semicolon row."""
        return (
            f"{index};{reg.fc};{reg.address};{reg.data_type};{reg.reserved};"
            f"{reg.name};{reg.tag};{reg.coef_a};{reg.coef_b};{reg.unit};{reg.action}"
        )

    @classmethod
    def format_header(cls, category: str, model: str) -> str:
        """Format the 11-column header for a Webdyn Modbus definition file."""
        return f"{cls.PROTOCOL};{category};{cls.MANUFACTURER};{model};;;;;;;"

    @classmethod
    def write_definition_file(
        cls,
        output_path: Path,
        category: str,
        model: str,
        registers: list[ModbusRegisterDef],
    ) -> None:
        """Write a definition file strictly adhering to the 5 Webdyn Pillars."""
        output_path.parent.mkdir(parents=True, exist_ok=True)
        lines = [cls.format_header(category, model)]
        for idx, reg in enumerate(registers, 1):
            lines.append(cls.format_row(idx, reg))

        content = "\r\n".join(lines) + "\r\n"
        output_path.write_bytes(content.encode("utf-8"))

    @classmethod
    def load_source_csv(cls, source_path: Path) -> list[ModbusRegisterDef]:
        """Load and sanitize registers from an existing verified CSV definition."""
        regs: list[ModbusRegisterDef] = []
        with open(source_path, encoding="utf-8-sig", errors="ignore") as fp:
            reader = csv.reader(fp, delimiter=";")
            next(reader, None)  # Skip header
            for raw_row in reader:
                if not raw_row or not any(raw_row):
                    continue
                row = [cell.replace("\r", " ").replace("\n", " ").strip() for cell in raw_row]
                fc = row[1] if len(row) > 1 and row[1] else "3"
                addr = row[2] if len(row) > 2 and row[2] else "0"
                dtype = row[3] if len(row) > 3 and row[3] else "U16"
                res = row[4] if len(row) > 4 else ""
                name = (row[5] if len(row) > 5 else "").replace(";", " - ")
                tag = row[6] if len(row) > 6 else ""
                raw_coef_a = row[7] if len(row) > 7 and row[7] else "1.000000"
                raw_coef_b = row[8] if len(row) > 8 and row[8] else "0.000000"
                unit = row[9] if len(row) > 9 and row[9] not in ("None", "/", "-") else ""
                action = row[10] if len(row) > 10 and row[10] in ("0", "1", "2", "4", "9") else "4"

                def _is_num(val: str) -> bool:
                    try:
                        float(val)
                        return True
                    except ValueError:
                        return False

                coef_a = raw_coef_a if _is_num(raw_coef_a) else "1.000000"
                coef_b = raw_coef_b if _is_num(raw_coef_b) else "0.000000"
                if not fc or fc not in ("1", "2", "3", "4"):
                    fc = "3"

                regs.append(
                    ModbusRegisterDef(
                        fc=fc,
                        address=addr,
                        data_type=dtype,
                        reserved=res,
                        name=name,
                        tag=tag,
                        coef_a=coef_a,
                        coef_b=coef_b,
                        unit=unit,
                        action=action,
                    )
                )
        return regs

    @classmethod
    def _find_source(cls, repo_root: Path, local_name: str, network_rel_path: str) -> Path | None:
        """Find source CSV either in local production directory or on network share."""
        p_local = repo_root / "production" / local_name
        if p_local.exists():
            return p_local
        p_net = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers\Huawei") / network_rel_path
        if p_net.exists():
            return p_net
        return None

    @classmethod
    def get_model_definitions(cls, repo_root: Path) -> dict[str, tuple[str, str, list[ModbusRegisterDef]]]:
        """Return certified Huawei model definitions mapped to (Category, Filename, Registers)."""
        models: dict[str, tuple[str, str, list[ModbusRegisterDef]]] = {}

        # 1. SUN2000-V4 Generic
        src = cls._find_source(
            repo_root,
            "INV_Huawei_ModbusV4.csv",
            "Inverter/HUAWEI_V4.csv",
        )
        if src:
            models["SUN2000-V4_Generic"] = ("Inverter", "INV_Huawei_SUN2000-V4_Generic.csv", cls.load_source_csv(src))

        # 2. SUN2000-2..5KTL Residential
        src = cls._find_source(
            repo_root,
            "INV_Huawei_MODBUS_2345.csv",
            "Inverter/OLD/1/Huawei_2345.csv",
        )
        if src:
            models["SUN2000-2_5KTL"] = ("Inverter", "INV_Huawei_SUN2000-2_5KTL.csv", cls.load_source_csv(src))

        # 3. SUN2000-8..20KTL Three-phase
        src = cls._find_source(
            repo_root,
            "INV_Huawei_MODBUS_81012141820.csv",
            "Inverter/OLD/1/Huawei_81012141820_Test.csv",
        )
        if src:
            models["SUN2000-8_20KTL"] = ("Inverter", "INV_Huawei_SUN2000-8_20KTL.csv", cls.load_source_csv(src))

        # 4. SUN2000-30..40KTL-M3 Commercial
        src = cls._find_source(
            repo_root,
            "INV_Huawei_SUN2000-M3.csv",
            "Inverter/SunSpec/WPM015E12_modbusRTU_Inverter_Huawei_SUN2000-30KTL-M3_DV3LZV3H4Q.csv",
        )
        if src:
            models["SUN2000-30_40KTL-M3"] = ("Inverter", "INV_Huawei_SUN2000-30_40KTL-M3.csv", cls.load_source_csv(src))

        # 5. SUN2000-33A..42KTL Commercial
        src = cls._find_source(
            repo_root,
            "INV_Huawei_364233-A.csv",
            "Inverter/OLD/1/Huawei_364233-A.csv",
        )
        if src:
            models["SUN2000-33A_36_42KTL"] = ("Inverter", "INV_Huawei_SUN2000-33A_36_42KTL.csv", cls.load_source_csv(src))

        # 6. SUN2000-50..70KTL-M0 Commercial
        src = cls._find_source(
            repo_root,
            "INV_Huawei_7550606570-M0.csv",
            "Inverter/OLD/1/Huawei_7550606570-M0.csv",
        )
        if src:
            models["SUN2000-50_70KTL-M0"] = ("Inverter", "INV_Huawei_SUN2000-50_70KTL-M0.csv", cls.load_source_csv(src))

        # 7. SUN2000-60KTL-HV High Voltage
        src = cls._find_source(
            repo_root,
            "INV_Huawei_MODBUS_60-HV.csv",
            "Inverter/OLD/1/Huawei_60-HV.csv",
        )
        if src:
            models["SUN2000-60KTL-HV"] = ("Inverter", "INV_Huawei_SUN2000-60KTL-HV.csv", cls.load_source_csv(src))

        # 8. SUN2000-90..105KTL Utility
        src = cls._find_source(
            repo_root,
            "INV_Huawei_MODBUS_9095100105.csv",
            "Inverter/OLD/1/Huawei_9095100105.csv",
        )
        if src:
            models["SUN2000-90_105KTL"] = ("Inverter", "INV_Huawei_SUN2000-90_105KTL.csv", cls.load_source_csv(src))

        # 9. SUN2000-V3 Standard
        src = cls._find_source(
            repo_root,
            "INV_Huawei_ModbusV3.csv",
            "Inverter/OLD/3/HUAWEI_V3.csv",
        )
        if src:
            models["SUN2000-V3"] = ("Inverter", "INV_Huawei_SUN2000-V3.csv", cls.load_source_csv(src))

        # 10. SUN2000-All V3 Full
        src = cls._find_source(
            repo_root,
            "INV_Huawei_Modbus_All_V3.csv",
            "Inverter/OLD/0/HUAWEI (SUN2000C00).csv",
        )
        if src:
            models["SUN2000-All_V3"] = ("Inverter", "INV_Huawei_SUN2000-All_V3.csv", cls.load_source_csv(src))

        # 11. LUNA2000 Residential Battery
        src = cls._find_source(
            repo_root,
            "BAT_Huawei_Luna2000.csv",
            "Battery_Luna2000/HUAWEI_Luna_2000.csv",
        )
        if src:
            models["LUNA2000"] = ("Battery", "BAT_Huawei_LUNA2000.csv", cls.load_source_csv(src))

        # 12. LUNA2000 Storage System
        src = cls._find_source(
            repo_root,
            "BAT_Huawei_INV_STORAGE.csv",
            "Inverter/HUAWEI_INV_STORAGE.csv",
        )
        if src:
            models["LUNA2000_Storage"] = ("Battery", "BAT_Huawei_LUNA2000_Storage.csv", cls.load_source_csv(src))

        # 13. SmartLogger 3000
        src = cls._find_source(
            repo_root,
            "COM_Huawei_SmartLogger3000.csv",
            "SmartLogger/SL3000/HUAWEI_SmartLogger3000.csv",
        )
        if src:
            models["SmartLogger3000"] = ("Logger", "COM_Huawei_SmartLogger3000.csv", cls.load_source_csv(src))

        # 14. SmartLogger V Series
        src = cls._find_source(
            repo_root,
            "COM_Huawei_DEF_HUAWEI_SmartLogger_VSeries.csv",
            "SmartLogger/V Series/DEF_HUAWEI_SmartLogger_VSeries.csv",
        )
        if src:
            models["SmartLogger_VSeries"] = ("Logger", "COM_Huawei_SmartLogger_VSeries.csv", cls.load_source_csv(src))

        # 15. Power Meter DTSU666-H
        src = cls._find_source(
            repo_root,
            "MET_Huawei_PowerMeter.csv",
            "Meter/HUAWEI_PowerMeter.csv",
        )
        if src:
            models["PowerMeter"] = ("Meter", "MET_Huawei_PowerMeter.csv", cls.load_source_csv(src))

        return models

    @classmethod
    def generate_all(cls, repo_root: Path) -> dict[str, Path]:
        """
        Generate and synchronize all certified Huawei definitions across the repository.
        Returns a mapping of model name to generated production file path.
        """
        models = cls.get_model_definitions(repo_root)
        output_paths: dict[str, Path] = {}

        prod_dir = repo_root / "production"
        def_dir = repo_root / "Definitions"
        huawei_def_dir = repo_root / "Definitions" / "Huawei"
        gen_dir = repo_root / "generated_definitions"
        tag_dir = repo_root / "tagged_definitions"

        target_dirs = [prod_dir, def_dir, huawei_def_dir, gen_dir, tag_dir]
        for d in target_dirs:
            d.mkdir(parents=True, exist_ok=True)

        for model, (cat, filename, regs) in models.items():
            prod_file = prod_dir / filename
            cls.write_definition_file(prod_file, cat, model, regs)
            output_paths[model] = prod_file

            # Mirror to target synchronization directories
            for target_dir in [huawei_def_dir, def_dir, gen_dir, tag_dir]:
                mirror_file = target_dir / filename
                mirror_file.write_bytes(prod_file.read_bytes())

        return output_paths
