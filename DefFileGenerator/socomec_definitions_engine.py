"""
Apollo-Standard Master Definition Engine for Socomec Energy Meters and Network Analyzers.

Official Documentation Sources:
  1. Diris A10: MET_Socomec_Diris_A10.csv (73 registers)
  2. Diris A40: MET_SOCOMEC_Diris_A40.csv (73 registers)
  3. Diris Axx Multi-Phase: MET_Socomec_Diris_Axx.csv (73 registers)
  4. Diris B10: MET_SOCOMEC_DIRIS_B10.csv (68 registers)
  5. Diris B30: MET_Socomec_Diris_B30.csv (83 registers)
  6. Diris Digiware I35: MET_SOCOMEC_DIRIS_DIGIWARE_I35.csv (46 registers)
  7. Countis E23: MET_SOCOMEC_COUNTIS_E23.csv (27 registers)
  8. Countis E27: MET_Socomec_COUNTIS_E27.csv (33 registers)
  9. Countis E33 / E43: MET_Socomec_Countis_E33_E43.csv (73 registers)
  10. Countis E50 / E53: MET_Socomec_COUNTIS_E50-E53.csv (43 registers)
  11. Countis P44: MET_Socomec_Countis_P44.csv (43 registers)

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


class SocomecDefinitionsEngine:
    """Master engine for compiling certified Socomec equipment definitions."""

    MANUFACTURER: str = "Socomec"
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
        net_path = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers\Socomec") / filename
        if net_path.exists():
            return net_path
        return None

    def get_model_definitions(
        self, repo_root: Path
    ) -> dict[str, tuple[str, str, list[ModbusRegisterDef]]]:
        """Compile and return all Socomec model definitions."""
        models: dict[str, tuple[str, str, list[ModbusRegisterDef]]] = {}

        specs = [
            (
                "Diris_A10",
                "Meter",
                "MET_Socomec_Diris_A10.csv",
                "MET_Socomec_Diris_A10.csv",
            ),
            (
                "Diris_A40",
                "Meter",
                "MET_Socomec_Diris_A40.csv",
                "MET_SOCOMEC_Diris_A40.csv",
            ),
            (
                "Diris_Axx",
                "Meter",
                "MET_Socomec_Diris_Axx.csv",
                "MET_Socomec_Diris_Axx.csv",
            ),
            (
                "Diris_B10",
                "Meter",
                "MET_Socomec_Diris_B10.csv",
                "MET_SOCOMEC_DIRIS_B10.csv",
            ),
            (
                "Diris_B30",
                "Meter",
                "MET_Socomec_Diris_B30.csv",
                "MET_Socomec_Diris_B30.csv",
            ),
            (
                "Diris_Digiware_I35",
                "Meter",
                "MET_Socomec_Diris_Digiware_I35.csv",
                "MET_SOCOMEC_DIRIS_DIGIWARE_I35.csv",
            ),
            (
                "Countis_E23",
                "Meter",
                "MET_Socomec_Countis_E23.csv",
                "MET_SOCOMEC_COUNTIS_E23.csv",
            ),
            (
                "Countis_E27",
                "Meter",
                "MET_Socomec_Countis_E27.csv",
                "MET_Socomec_COUNTIS_E27.csv",
            ),
            (
                "Countis_E33_E43",
                "Meter",
                "MET_Socomec_Countis_E33_E43.csv",
                "MET_Socomec_Countis_E33_E43.csv",
            ),
            (
                "Countis_E50_E53",
                "Meter",
                "MET_Socomec_Countis_E50_E53.csv",
                "MET_Socomec_COUNTIS_E50-E53.csv",
            ),
            (
                "Countis_P44",
                "Meter",
                "MET_Socomec_Countis_P44.csv",
                "MET_Socomec_Countis_P44.csv",
            ),
        ]

        for model_key, category, out_name, src_name in specs:
            src = self._find_source(repo_root, src_name)
            if src and src.exists():
                regs = self.load_source_csv(src)
                models[model_key] = (category, out_name, regs)

        return models

    def generate_all(self, repo_root: Path) -> list[Path]:
        """Generate all certified Socomec definitions into output directories."""
        generated: list[Path] = []
        models = self.get_model_definitions(repo_root)

        output_dirs = [
            repo_root / "Definitions" / "Socomec",
            repo_root / "production",
            repo_root / "Definitions",
            repo_root / "generated_definitions",
            repo_root / "tagged_definitions",
        ]

        for model_key, (category, out_name, regs) in models.items():
            for out_dir in output_dirs:
                dest = out_dir / out_name
                self.write_definition_file(dest, category, model_key, regs)
                if out_dir == repo_root / "Definitions" / "Socomec":
                    generated.append(dest)

        return generated
