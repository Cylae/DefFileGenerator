"""Filesystem safeguards shared by the command-line and desktop interfaces."""

import csv
import io
import os
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from typing import TextIO


def paths_refer_to_same_file(source: str, destination: str) -> bool:
    """Recognize identical paths, symbolic links, and existing hard-link aliases."""
    if os.path.normcase(os.path.realpath(source)) == os.path.normcase(
        os.path.realpath(destination)
    ):
        return True
    try:
        return os.path.samefile(source, destination)
    except OSError:
        return False


def _get_field(row: list[str], idx: int, default: str = "") -> str:
    return row[idx] if idx < len(row) else default


def definition_preview(csv_content: str, limit: int) -> list[dict[str, str]]:
    """Read a preview from a validated definition's actual serialized fields."""
    reader = csv.reader(io.StringIO(csv_content), delimiter=";")
    next(reader, None)
    register_types = {
        "1": "Coils",
        "2": "Discrete Inputs",
        "3": "Holding Register",
        "4": "Input Register",
    }
    rows = []
    for index, row in enumerate(reader):
        if index >= limit:
            break
        if not row:
            continue
        reg_type_raw = _get_field(row, 1)
        rows.append(
            {
                "RegisterType": register_types.get(reg_type_raw, reg_type_raw),
                "Address": _get_field(row, 2),
                "Type": _get_field(row, 3),
                "Name": _get_field(row, 5),
                "Tag": _get_field(row, 6),
                "Factor": _get_field(row, 7),
                "Offset": _get_field(row, 8),
                "Unit": _get_field(row, 9),
                "Action": _get_field(row, 10),
            }
        )
    return rows


@contextmanager
def staged_text_output(destination: str, encoding: str = "utf-8") -> Iterator[TextIO]:
    """Replace a text file only after the complete write and disk sync succeed."""
    target = os.path.abspath(destination)
    descriptor, staging_path = tempfile.mkstemp(
        prefix=f".{os.path.basename(target)}.", suffix=".tmp", dir=os.path.dirname(target)
    )
    try:
        with os.fdopen(descriptor, "w", newline="", encoding=encoding) as output:
            yield output
            output.flush()
            os.fsync(output.fileno())
        os.replace(staging_path, target)
    finally:
        if os.path.exists(staging_path):
            os.unlink(staging_path)
