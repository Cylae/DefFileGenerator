"""Regression coverage for installed API and legacy CLI failure boundaries."""

from pathlib import Path
from unittest.mock import patch

import pytest

import doc_to_webdyn
import generate_webdyn_def


def _input(path: Path) -> Path:
    path.write_text("Address,Name,Type\n100,Power,U16\n", encoding="utf-8")
    return path


def test_api_catches_deferred_extraction_errors(tmp_path: Path) -> None:
    source = _input(tmp_path / "input.csv")
    output = tmp_path / "output.csv"
    output.write_bytes(b"existing definition")

    def broken_stream():
        raise RuntimeError("deferred parser failure")
        yield  # pragma: no cover

    with patch.object(
        generate_webdyn_def.Extractor, "extract_from_csv", return_value=broken_stream()
    ):
        assert not generate_webdyn_def.generate_webdyn_definition(
            str(source), str(output), "Vendor", "Device"
        )
    assert output.read_bytes() == b"existing definition"


def test_api_does_not_validate_stale_output_after_failed_generation(tmp_path: Path) -> None:
    source = _input(tmp_path / "input.csv")
    output = tmp_path / "output.csv"
    output.write_bytes(b"existing definition")
    with (
        patch.object(generate_webdyn_def, "run_generator", return_value=False),
        patch.object(generate_webdyn_def.Generator, "validate_csv", return_value=True) as validate,
    ):
        assert not generate_webdyn_def.generate_webdyn_definition(
            str(source), str(output), "Vendor", "Device"
        )
        validate.assert_not_called()
    assert output.read_bytes() == b"existing definition"


@pytest.mark.parametrize("alias", [False, True])
def test_api_preserves_source_when_output_is_same_file(tmp_path: Path, alias: bool) -> None:
    source = _input(tmp_path / "input.csv")
    before = source.read_bytes()
    output = source
    if alias:
        output = tmp_path / "alias.csv"
        try:
            output.hardlink_to(source)
        except OSError as exc:
            pytest.skip(f"Hard links unavailable: {exc}")
    assert not generate_webdyn_def.generate_webdyn_definition(
        str(source), str(output), "Vendor", "Device"
    )
    assert source.read_bytes() == before
    assert output.read_bytes() == before


def test_legacy_cli_reports_failed_generation(tmp_path: Path) -> None:
    source = _input(tmp_path / "input.csv")
    with patch.object(doc_to_webdyn, "run_generator", return_value=False):
        with pytest.raises(SystemExit) as error:
            doc_to_webdyn.main([str(source), "-o", str(tmp_path / "output.csv")])
    assert error.value.code == 1


def test_legacy_cli_preserves_existing_output_without_force(tmp_path: Path) -> None:
    source = _input(tmp_path / "input.csv")
    output = tmp_path / "output.csv"
    output.write_bytes(b"existing definition")
    with pytest.raises(SystemExit) as error:
        doc_to_webdyn.main([str(source), "-o", str(output)])
    assert error.value.code == 1
    assert output.read_bytes() == b"existing definition"


def test_legacy_cli_force_replaces_distinct_output(tmp_path: Path) -> None:
    source = _input(tmp_path / "input.csv")
    output = tmp_path / "output.csv"
    output.write_bytes(b"existing definition")
    doc_to_webdyn.main([str(source), "-o", str(output), "--force"])
    assert ";100;U16;" in output.read_text(encoding="utf-8-sig")


def test_legacy_cli_never_overwrites_source_even_with_force(tmp_path: Path) -> None:
    source = _input(tmp_path / "input.csv")
    before = source.read_bytes()
    with pytest.raises(SystemExit) as error:
        doc_to_webdyn.main([str(source), "-o", str(source), "--force"])
    assert error.value.code == 1
    assert source.read_bytes() == before
