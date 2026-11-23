"""
Comprehensive Apollo-Standard Test Suite for SolarEdge Definitions Engine.
"""

from pathlib import Path
import pytest

from DefFileGenerator.solaredge_definitions_engine import SolarEdgeDefinitionsEngine

REPO_ROOT = Path(__file__).resolve().parent.parent.parent


@pytest.fixture(scope="module")
def solaredge_engine() -> SolarEdgeDefinitionsEngine:
    return SolarEdgeDefinitionsEngine()


@pytest.fixture(scope="module")
def solaredge_models(solaredge_engine: SolarEdgeDefinitionsEngine) -> dict:
    return solaredge_engine.get_model_definitions(REPO_ROOT)


def test_solaredge_models_coverage(solaredge_models: dict) -> None:
    """Verify that all core SolarEdge models are discovered and loaded."""
    expected_models = {
        "SE33.3K_SunSpec",
        "SE25K_SunSpec",
        "SunSpec_Generic",
        "66KW",
        "EnhancedPowerControl",
        "ActivePowerControl",
        "SE_Series",
        "SE-MTR-3Y-400V-A",
    }
    assert expected_models.issubset(set(solaredge_models.keys()))


def test_solaredge_5_pillars_certification(solaredge_models: dict, tmp_path: Path) -> None:
    """Verify that each SolarEdge model satisfies the 5 Apollo-standard pillars."""
    engine = SolarEdgeDefinitionsEngine()
    for model_key, (category, out_filename, regs) in solaredge_models.items():
        assert len(regs) > 0, f"Model {model_key} has 0 registers"
        test_file = tmp_path / out_filename
        engine.write_definition_file(test_file, category, model_key, regs)

        raw = test_file.read_bytes()
        # Pillar 1 & 2: UTF-8 without BOM, CRLF
        assert not raw.startswith(b"\xef\xbb\xbf"), f"BOM detected in {out_filename}"
        assert raw.endswith(b"\r\n"), f"Missing CRLF in {out_filename}"
        assert not raw.endswith(b"\r\r\n"), f"Double CR detected in {out_filename}"
        stripped = raw.replace(b"\r\n", b"")
        assert b"\n" not in stripped, f"Isolated LF in {out_filename}"
        assert b"\r" not in stripped, f"Isolated CR in {out_filename}"

        # Pillar 3: 11 columns
        lines = raw.decode("utf-8").splitlines()
        header = lines[0].split(";")
        assert len(header) == 11, f"Header of {out_filename} has {len(header)} cols"
        assert header[0] == "Modbus"
        assert header[2] == "SolarEdge"

        for idx, line in enumerate(lines[1:], 1):
            cols = line.split(";")
            assert len(cols) == 11, f"Row {idx} in {out_filename} has {len(cols)} cols"
            assert cols[0] == str(idx), f"Index mismatch at row {idx}"
            assert cols[1] in ("1", "2", "3", "4"), f"Invalid FC {cols[1]} at row {idx}"
            assert cols[10] in ("0", "1", "2", "4", "9"), f"Invalid Action {cols[10]} at row {idx}"


def test_solaredge_generation(solaredge_engine: SolarEdgeDefinitionsEngine) -> None:
    """Verify that generation populates all mirror directories without errors."""
    generated = solaredge_engine.generate_all(REPO_ROOT)
    assert len(generated) >= 8
    for p in generated:
        assert p.exists()
        assert p.stat().st_size > 0
