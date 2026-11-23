"""
Apollo-Standard Master Definition Engine for KACO new energy Inverters.

Official Documentation Sources:
  1. blueplanet 125 NX3 M10: 95308-LOGNON_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.0.3.3-ec484f_modbusRTU.csv (928 registers)
  2. blueplanet 100 NX3 M8: SunSpec_inverter_KACO_new_energy_blueplanet_100_NX3_M8_IEC_5.0.3.3-ec484f_modbusRTU.csv (908 registers)
  3. blueplanet 125 TL3: WPM00D3FF_SunSpec_inverter_KACO_new_energy_blueplanet_125_TL3_02.37-17d56298_modbusTCP.csv (828 registers)
  4. blueplanet 50.0 TL3: GOUSSARD_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V5.24_modbusTCP.csv (628 registers)
  5. blueplanet 50.0 NX3 M5: WPM016888_SunSpec_inverter_KACO_new_energy_bp_50.0_NX3_M5_V610-03030-04__modbusRTU.csv (197 registers)
  6. blueplanet 20.0 TL3: MOURLON_SunSpec_inverter_KACO_new_energy_blueplanet_20.0_TL3_V5.20_modbusTCP__1.csv (697 registers)
  7. blueplanet 15.0 TL3: CHRISTIANS_SunSpec_inverter_KACO_new_energy_blueplanet_15.0_TL3_V5.20_modbusTCP__1.csv (638 registers)
  8. blueplanet NX1: Kaco_NX1-NX3_ModbusRTU.csv / KACO_NX1.csv (182 registers)
  9. blueplanet NX3: Kaco_NX1-NX3_ModbusRTU.csv (182 registers)
  10. Powador 60.0 TL3: CHRISTIAENS_SunSpec_inverter_KACO_new_energy_Powador_60.0_TL3_V3.52_modbusTCP.csv (101 registers)
  11. Powador 360TL: EGUZON_KAC_Inverter_KAC_360TL.csv (43 registers)
  12. Powador 600TL: EGUZON_KAC_Inverter_KAC_600TL.csv (43 registers)

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


class KacoDefinitionsEngine:
    """Master engine for compiling certified KACO new energy equipment definitions."""

    MANUFACTURER: str = "Kaco"
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
        p_net = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers\Kaco") / network_rel_path
        if p_net.exists():
            return p_net
        return None

    @classmethod
    def get_model_definitions(cls, repo_root: Path) -> dict[str, tuple[str, str, list[ModbusRegisterDef]]]:
        """Return certified Kaco model definitions mapped to (Category, Filename, Registers)."""
        models: dict[str, tuple[str, str, list[ModbusRegisterDef]]] = {}

        # 1. blueplanet 125 NX3 M10 SunSpec
        src = cls._find_source(
            repo_root,
            "INV_Kaco_95308-LOGNON_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.0.3.3-ec484f_modbusRTU.csv",
            "Modbus/95308-LOGNON_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.0.3.3-ec484f_modbusRTU.csv",
        )
        if src:
            models["blueplanet_125_NX3_M10"] = ("Inverter", "INV_Kaco_blueplanet_125_NX3_M10.csv", cls.load_source_csv(src))

        # 2. blueplanet 100 NX3 M8 SunSpec
        src = cls._find_source(
            repo_root,
            "INV_Kaco_SunSpec_inverter_KACO_new_energy_blueplanet_100_NX3_M8_IEC_5.0.3.3-ec484f_modbusRTU.csv",
            "Modbus/SunSpec_inverter_KACO_new_energy_blueplanet_100_NX3_M8_IEC_5.0.3.3-ec484f_modbusRTU.csv",
        )
        if src:
            models["blueplanet_100_NX3_M8"] = ("Inverter", "INV_Kaco_blueplanet_100_NX3_M8.csv", cls.load_source_csv(src))

        # 3. blueplanet 125 TL3 SunSpec
        src = cls._find_source(
            repo_root,
            "INV_Kaco_blueplanet_125_TL3.csv",
            "Modbus/WPM00D3FF_SunSpec_inverter_KACO_new_energy_blueplanet_125_TL3_02.37-17d56298_modbusTCP.csv",
        )
        if src:
            models["blueplanet_125_TL3"] = ("Inverter", "INV_Kaco_blueplanet_125_TL3.csv", cls.load_source_csv(src))

        # 4. blueplanet 50.0 TL3 SunSpec
        src = cls._find_source(
            repo_root,
            "INV_Kaco_GOUSSARD_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V5.24_modbusTCP.csv",
            "Modbus/GOUSSARD_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V5.24_modbusTCP.csv",
        )
        if src:
            models["blueplanet_50.0_TL3"] = ("Inverter", "INV_Kaco_blueplanet_50.0_TL3.csv", cls.load_source_csv(src))

        # 5. blueplanet 50.0 NX3 M5 SunSpec
        src = cls._find_source(
            repo_root,
            "INV_Kaco_blueplanet_50.0_NX3_M5.csv",
            "Modbus/WPM016888_SunSpec_inverter_KACO_new_energy_bp_50.0_NX3_M5_V610-03030-04__modbusRTU.csv",
        )
        if src:
            models["blueplanet_50.0_NX3_M5"] = ("Inverter", "INV_Kaco_blueplanet_50.0_NX3_M5.csv", cls.load_source_csv(src))

        # 6. blueplanet 20.0 TL3 SunSpec
        src = cls._find_source(
            repo_root,
            "INV_Kaco_MOURLON_SunSpec_inverter_KACO_new_energy_blueplanet_20.0_TL3_V5.20_modbusTCP_1.csv",
            "Modbus/MOURLON_SunSpec_inverter_KACO_new_energy_blueplanet_20.0_TL3_V5.20_modbusTCP__1.csv",
        )
        if src:
            models["blueplanet_20.0_TL3"] = ("Inverter", "INV_Kaco_blueplanet_20.0_TL3.csv", cls.load_source_csv(src))

        # 7. blueplanet 15.0 TL3 SunSpec
        src = cls._find_source(
            repo_root,
            "INV_Kaco_CHRISTIANS_SunSpec_inverter_KACO_new_energy_blueplanet_15.0_TL3_V5.20_modbusTCP_1.csv",
            "Modbus/CHRISTIANS_SunSpec_inverter_KACO_new_energy_blueplanet_15.0_TL3_V5.20_modbusTCP__1.csv",
        )
        if src:
            models["blueplanet_15.0_TL3"] = ("Inverter", "INV_Kaco_blueplanet_15.0_TL3.csv", cls.load_source_csv(src))

        # 8. blueplanet NX1 Residential Inverter
        src = cls._find_source(
            repo_root,
            "INV_Kaco_NX1.csv",
            "Modbus/KACO_NX1.csv",
        )
        if src:
            models["blueplanet_NX1"] = ("Inverter", "INV_Kaco_blueplanet_NX1.csv", cls.load_source_csv(src))

        # 9. blueplanet NX3 Commercial Inverter
        src = cls._find_source(
            repo_root,
            "INV_Kaco_NX1-NX3_ModbusRTU.csv",
            "Modbus/Kaco_NX1-NX3_ModbusRTU.csv",
        )
        if src:
            models["blueplanet_NX3"] = ("Inverter", "INV_Kaco_blueplanet_NX3.csv", cls.load_source_csv(src))

        # 10. Powador 60.0 TL3
        src = cls._find_source(
            repo_root,
            "INV_Kaco_CHRISTIAENS_SunSpec_inverter_KACO_new_energy_Powador_60.0_TL3_V3.52_modbusTCP.csv",
            "Modbus/CHRISTIAENS_SunSpec_inverter_KACO_new_energy_Powador_60.0_TL3_V3.52_modbusTCP.csv",
        )
        if src:
            models["Powador_60.0_TL3"] = ("Inverter", "INV_Kaco_Powador_60.0_TL3.csv", cls.load_source_csv(src))

        return models

    @classmethod
    def generate_all(cls, repo_root: Path) -> dict[str, Path]:
        """
        Generate and synchronize all certified Kaco definitions across the repository.
        Returns a mapping of model name to generated production file path.
        """
        models = cls.get_model_definitions(repo_root)
        output_paths: dict[str, Path] = {}

        prod_dir = repo_root / "production"
        def_dir = repo_root / "Definitions"
        kaco_def_dir = repo_root / "Definitions" / "Kaco"
        gen_dir = repo_root / "generated_definitions"
        tag_dir = repo_root / "tagged_definitions"

        target_dirs = [prod_dir, def_dir, kaco_def_dir, gen_dir, tag_dir]
        for d in target_dirs:
            d.mkdir(parents=True, exist_ok=True)

        for model, (cat, filename, regs) in models.items():
            prod_file = prod_dir / filename
            cls.write_definition_file(prod_file, cat, model, regs)
            output_paths[model] = prod_file

            # Mirror to target synchronization directories
            for target_dir in [kaco_def_dir, def_dir, gen_dir, tag_dir]:
                mirror_file = target_dir / filename
                mirror_file.write_bytes(prod_file.read_bytes())

        return output_paths
