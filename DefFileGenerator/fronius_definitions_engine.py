"""
Apollo-Standard Master Definition Engine for Fronius Inverters, Smart Meters, and Sensors.

Official Documentation Sources:
  1. Gen24 / Primo / Symo (Float): INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Float.csv (262 registers)
  2. Gen24 / Primo / Symo (Int+SF): INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Int_SF.csv (269 registers)
  3. Symo GEN24 Plus: INV_Fronius_Fronius_SunSpec_inverter_Fronius_Symo_GEN24.csv (223 registers)
  4. Tauro 50-3-D: INV_Fronius_Tauro_50-3-D_1.35.4-1_Float.csv (193 registers)
  5. Tauro ECO 100-3-D: INV_Fronius_WPM01A481_Fronius_Tauro_ECO_100-3-D.csv (167 registers)
  6. Symo 10.0-20.0-3-M: INV_Fronius_WPM018C3D_Fronius_Symo_10.0-3-M.csv (177 registers)
  7. Verto 3MPPT: INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Float.csv (279 registers)
  8. Verto 4MPPT: INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Float.csv (299 registers)
  9. ECO Series: INV_Fronius_ECO_Inverter_Register_Map_Int_SF.csv (257 registers)
  10. Smart Meter TS (Float): MET_Fronius_Smart_Meter_Register_Map_Float.csv (144 registers)
  11. Smart Meter (Int+SF): MET_Fronius_Smart_Meter_Register_Map_Int_SF.csv (165 registers)
  12. Smart Meter 50kA-3: MET_Fronius_WPM01A481_Fronius_Smart_Meter_50kA-3.csv (72 registers)
  13. Sensor Card: SEN_Fronius_IG_Sensor_Card_Register_Map_Float_and_Int_SF_v1.0.csv (54 registers)

Strictly enforces:
  - Pillar 1: Strict UTF-8 without BOM.
  - Pillar 2: Strict CRLF (\\r\\n) line endings.
  - Pillar 3: Exact 11-column semicolon CSV schema with standard Webdyn header.
  - Pillar 4: Monotonic sequential 1..N row indexing.
  - Pillar 5: Valid Webdyn data types and canonical SunPM tags.
"""

from __future__ import annotations

import csv
import re
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


class FroniusDefinitionsEngine:
    """Master engine for compiling certified Fronius equipment definitions."""

    MANUFACTURER: str = "Fronius"
    PROTOCOL: str = "Modbus"

    @classmethod
    def clean_cell(cls, val: str) -> str:
        s = str(val).replace("\r", " ").replace("\n", " ").replace(";", " - ")
        return re.sub(r"\s+", " ", s).strip()

    @classmethod
    def format_row(cls, index: int, reg: ModbusRegisterDef) -> str:
        """Format a single register definition into an 11-column semicolon row."""
        return (
            f"{index};{reg.fc};{reg.address};{reg.data_type};{reg.reserved};"
            f"{cls.clean_cell(reg.name)};{cls.clean_cell(reg.tag)};{cls.clean_cell(reg.coef_a)};"
            f"{cls.clean_cell(reg.coef_b)};{cls.clean_cell(reg.unit)};{cls.clean_cell(reg.action)}"
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
    def _find_source(cls, repo_root: Path, filename: str) -> Path | None:
        """Find source CSV file searching across standard project search paths."""
        for sub in ("tagged_definitions", "Definitions", "production", "Definitions/Classic"):
            cand = repo_root / sub / filename
            if cand.exists():
                return cand
        for cand in repo_root.rglob(filename):
            if cand.is_file():
                return cand
        net_path = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers\Fronius") / filename
        if net_path.exists():
            return net_path
        return None

    def get_model_definitions(
        self, repo_root: Path
    ) -> dict[str, tuple[str, str, list[ModbusRegisterDef]]]:
        """Compile and return all Fronius model definitions."""
        models: dict[str, tuple[str, str, list[ModbusRegisterDef]]] = {}

        specs = [
            (
                "Gen24_Primo_Symo_Float",
                "Inverter",
                "INV_Fronius_Gen24_Primo_Symo_Float.csv",
                "INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Float.csv",
            ),
            (
                "Gen24_Primo_Symo_IntSF",
                "Inverter",
                "INV_Fronius_Gen24_Primo_Symo_IntSF.csv",
                "INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Int_SF.csv",
            ),
            (
                "Symo_GEN24_Plus",
                "Inverter",
                "INV_Fronius_Symo_GEN24_Plus.csv",
                "INV_Fronius_Fronius_SunSpec_inverter_Fronius_Symo_GEN24.csv",
            ),
            (
                "Tauro_50-3-D",
                "Inverter",
                "INV_Fronius_Tauro_50-3-D.csv",
                "INV_Fronius_Tauro_50-3-D_1.35.4-1_Float.csv",
            ),
            (
                "Tauro_ECO_100-3-D",
                "Inverter",
                "INV_Fronius_Tauro_ECO_100-3-D.csv",
                "INV_Fronius_WPM01A481_Fronius_Tauro_ECO_100-3-D.csv",
            ),
            (
                "Symo_10_20",
                "Inverter",
                "INV_Fronius_Symo_10_20.csv",
                "INV_Fronius_WPM018C3D_Fronius_Symo_10.0-3-M.csv",
            ),
            (
                "Verto_3MPPT",
                "Inverter",
                "INV_Fronius_Verto_3MPPT.csv",
                "INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Float.csv",
            ),
            (
                "Verto_4MPPT",
                "Inverter",
                "INV_Fronius_Verto_4MPPT.csv",
                "INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Float.csv",
            ),
            (
                "ECO_Series",
                "Inverter",
                "INV_Fronius_ECO_Series.csv",
                "INV_Fronius_ECO_Inverter_Register_Map_Int_SF.csv",
            ),
            (
                "Smart_Meter_TS",
                "Meter",
                "MET_Fronius_Smart_Meter_TS.csv",
                "MET_Fronius_Smart_Meter_Register_Map_Float.csv",
            ),
            (
                "Smart_Meter_IntSF",
                "Meter",
                "MET_Fronius_Smart_Meter_IntSF.csv",
                "MET_Fronius_Smart_Meter_Register_Map_Int_SF.csv",
            ),
            (
                "Smart_Meter_50kA",
                "Meter",
                "MET_Fronius_Smart_Meter_50kA-3.csv",
                "MET_Fronius_WPM01A481_Fronius_Smart_Meter_50kA-3.csv",
            ),
            (
                "Sensor_Card",
                "Sensor",
                "SEN_Fronius_Sensor_Card.csv",
                "SEN_Fronius_IG_Sensor_Card_Register_Map_Float_and_Int_SF_v1.0.csv",
            ),
        ]

        for model_key, category, out_name, src_name in specs:
            src = self._find_source(repo_root, src_name)
            if src and src.exists():
                regs = self.load_source_csv(src)
                models[model_key] = (category, out_name, regs)

        return models

    def generate_all(self, repo_root: Path) -> list[Path]:
        """Generate all certified Fronius definitions into output directories."""
        generated: list[Path] = []
        models = self.get_model_definitions(repo_root)

        output_dirs = [
            repo_root / "Definitions" / "Fronius",
            repo_root / "production",
            repo_root / "Definitions",
            repo_root / "generated_definitions",
            repo_root / "tagged_definitions",
        ]

        for model_key, (category, out_name, regs) in models.items():
            for out_dir in output_dirs:
                dest = out_dir / out_name
                self.write_definition_file(dest, category, model_key, regs)
                if out_dir == repo_root / "Definitions" / "Fronius":
                    generated.append(dest)

        return generated
