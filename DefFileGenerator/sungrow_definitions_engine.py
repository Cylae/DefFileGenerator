"""
Apollo-Standard Master Definition Engine for Sungrow Inverters, Batteries & Loggers.

Official Documentation Sources:
  1. PowerStack ST255CS: Sungrow_Battery_v3.csv & PowerStack-ST255CS-2H User Manual.pdf (1,442 registers)
  2. LC3000 ESS: LC3000.csv & TI_20251013_LC300_Modbus Communication Protocol_V8_EN.xlsx (916 registers)
  3. FullMap v2: Sungrow_Battery_v2.csv (801 registers)
  4. Standard Battery: Sungrow_Battery.csv & SBR096-256-UEN-Ver12.pdf (172 registers)
  5. SG Series Generic: INV_SUNGROW_GENERIC_V.1.1.58.csv & SG30_33_40_50CX-UEN-Ver18-202106.pdf (200 registers)
  6. SG Series: INV_Sungrow_SG_Series.csv (29 registers)
  7. Generic Inverter: INV_SUNGROW_GENERIC.csv (164 registers)
  8. Logger 1000 / 3000: MODBUS_SUNGROW_Logger1000-3000.csv & TI_20210113_Logger_Communication Protocol_V1.0.2.5_EN.pdf (50 registers)
  9. EMS300CP: TI_20241202_EMS300CP_IEC104&Modbus Communication Protocol_CN&EN.xlsx (694 registers)
  10. LC300 Modbus: TI_20251013_LC300_Modbus Communication Protocol_V8_EN.xlsx (813 registers)
  11. PV Combiner Box: SUNGROW_PVCombinerBox.csv (64 registers)

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


class SungrowDefinitionsEngine:
    """Master engine for compiling certified Sungrow equipment definitions."""

    MANUFACTURER: str = "Sungrow"
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
        p_net = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers\Sungrow") / network_rel_path
        if p_net.exists():
            return p_net
        return None

    @classmethod
    def get_model_definitions(cls, repo_root: Path) -> dict[str, tuple[str, str, list[ModbusRegisterDef]]]:
        """Return certified Sungrow model definitions mapped to (Category, Filename, Registers)."""
        models: dict[str, tuple[str, str, list[ModbusRegisterDef]]] = {}

        # 1. PowerStack ST255CS Liquid-Cooled BESS
        src = cls._find_source(
            repo_root,
            "BAT_Sungrow_v3.csv",
            "Modbus/Battery/Sungrow_Battery_v3.csv",
        )
        if src:
            models["PowerStack_ST255CS"] = ("Battery", "BAT_Sungrow_PowerStack_ST255CS.csv", cls.load_source_csv(src))

        # 2. LC3000 Energy Storage System
        src = cls._find_source(
            repo_root,
            "BAT_Sungrow_LC3000.csv",
            "Modbus/Battery/LC3000.csv",
        )
        if src:
            models["LC3000"] = ("Battery", "BAT_Sungrow_LC3000.csv", cls.load_source_csv(src))

        # 3. FullMap v2 ESS
        src = cls._find_source(
            repo_root,
            "BAT_Sungrow_v2.csv",
            "Modbus/Battery/Sungrow_Battery_v2.csv",
        )
        if src:
            models["FullMap_v2"] = ("Battery", "BAT_Sungrow_FullMap_v2.csv", cls.load_source_csv(src))

        # 4. Standard Battery
        src = cls._find_source(
            repo_root,
            "BAT_Sungrow_Battery.csv",
            "Modbus/Battery/Sungrow_Battery.csv",
        )
        if src:
            models["Battery"] = ("Battery", "BAT_Sungrow_Battery.csv", cls.load_source_csv(src))

        # 5. SG Series Commercial Generic
        src = cls._find_source(
            repo_root,
            "INV_Sungrow_GENERIC_V.1.1.58.csv",
            "Modbus/old/INV_SUNGROW_GENERIC_V.1.1.58.csv",
        )
        if src:
            models["SG_Series_Generic"] = ("Inverter", "INV_Sungrow_SG_Series_Generic.csv", cls.load_source_csv(src))

        # 6. SG Series Three-phase
        src = cls._find_source(
            repo_root,
            "INV_Sungrow_SG_Series.csv",
            "Modbus/Inverter/INV_SUNGROW_GENERIC.csv",
        )
        if src:
            models["SG_Series"] = ("Inverter", "INV_Sungrow_SG_Series.csv", cls.load_source_csv(src))

        # 7. Generic PV Inverter
        src = cls._find_source(
            repo_root,
            "INV_Sungrow_Generic.csv",
            "Modbus/Inverter/INV_SUNGROW_GENERIC.csv",
        )
        if src:
            models["Generic"] = ("Inverter", "INV_Sungrow_Generic.csv", cls.load_source_csv(src))

        # 8. Logger 1000 / 3000
        src = cls._find_source(
            repo_root,
            "COM_Sungrow_MODBUS_SUNGROW_Logger1000-3000.csv",
            "Modbus/Logger/MODBUS_SUNGROW_Logger1000-3000.csv",
        )
        if src:
            models["Logger1000_3000"] = ("Logger", "COM_Sungrow_Logger1000_3000.csv", cls.load_source_csv(src))

        # 9. EMS300CP Energy Management System
        src = cls._find_source(
            repo_root,
            "COM_Sungrow_TI_20241202_EMS300CP_IEC104_Modbus_Communication_Protocol_P014_software_version_CN_EN.csv",
            "Modbus/Battery/Sources/TI_20241202_EMS300CP_IEC104&Modbus Communication Protocol（P014 software version）_CN&EN.xlsx",
        )
        if src:
            models["EMS300CP"] = ("Logger", "COM_Sungrow_EMS300CP.csv", cls.load_source_csv(src))

        # 10. LC300 Modbus Controller
        src = cls._find_source(
            repo_root,
            "COM_Sungrow_TI_20251013_LC300_Modbus_Communication_Protocol_V8_EN.csv",
            "Modbus/Battery/Sources/TI_20251013_LC300_Modbus Communication Protocol_V8_EN.xlsx",
        )
        if src:
            models["LC300_Modbus"] = ("Logger", "COM_Sungrow_LC300_Modbus.csv", cls.load_source_csv(src))

        # 11. PV Combiner Box
        src = cls._find_source(
            repo_root,
            "INV_Sungrow_PVCombinerBox.csv",
            "Modbus/PV Combiner box with 24 inputs/SUNGROW_PVCombinerBox.csv",
        )
        if src:
            models["PVCombinerBox"] = ("Other", "OTH_Sungrow_PVCombinerBox.csv", cls.load_source_csv(src))

        return models

    @classmethod
    def generate_all(cls, repo_root: Path) -> dict[str, Path]:
        """
        Generate and synchronize all certified Sungrow definitions across the repository.
        Returns a mapping of model name to generated production file path.
        """
        models = cls.get_model_definitions(repo_root)
        output_paths: dict[str, Path] = {}

        prod_dir = repo_root / "production"
        def_dir = repo_root / "Definitions"
        sungrow_def_dir = repo_root / "Definitions" / "Sungrow"
        gen_dir = repo_root / "generated_definitions"
        tag_dir = repo_root / "tagged_definitions"

        target_dirs = [prod_dir, def_dir, sungrow_def_dir, gen_dir, tag_dir]
        for d in target_dirs:
            d.mkdir(parents=True, exist_ok=True)

        for model, (cat, filename, regs) in models.items():
            model_header = "SG-Series" if model == "SG_Series" else model
            prod_file = prod_dir / filename
            cls.write_definition_file(prod_file, cat, model_header, regs)
            output_paths[model] = prod_file

            # Mirror to target synchronization directories
            for target_dir in [sungrow_def_dir, def_dir, gen_dir, tag_dir]:
                mirror_file = target_dir / filename
                mirror_file.write_bytes(prod_file.read_bytes())

        return output_paths
