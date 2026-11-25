"""
WebdynSun (Classic) Definition Exporter and Validator (.ini format).

Implements the official WebdynSun Gateway Specification (Version 2.0,
Sections 10.3-10.6):
  - `Modbus_RequestsTables={...}`: 10 semicolon-separated fields
    (reqIndex;reqName;reqReadFctCode;reqWriteFctCode;reqStartRegister;
     reqNbRegisters;reqEnableReading;reqEnableWriting;reqOption1;reqOption2)
  - `Modbus_VariablesTables={...}`: 12 semicolon-separated fields
    (varIndex;varReqIndex;varName;varType;varSigned;varPosition;
     varOption1;varOption2;coeffA;coeffB;unit;varAction)
"""

from __future__ import annotations

import csv
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from DefFileGenerator.io_utils import staged_text_output

# Maximum 16-bit registers per contiguous read request on WebdynSun
MAX_WEBDYNSUN_READ_REGS = 32

# Official WebdynSun varType mapping from WebdynSunPM Info3 codes
_SUNPM_TO_SUN_TYPE: dict[str, tuple[int, int, int]] = {
    # info3_upper: (varType, varSigned [1=signed, 2=unsigned], reg_count)
    "BITS": (1, 2, 1),
    "U8": (2, 2, 1),
    "U8_B": (2, 2, 1),
    "U8_W": (2, 2, 1),
    "U8_WB": (2, 2, 1),
    "I8": (2, 1, 1),
    "I8_B": (2, 1, 1),
    "I8_W": (2, 1, 1),
    "I8_WB": (2, 1, 1),
    "U16": (3, 2, 1),
    "U16_B": (4, 2, 1),
    "U16_W": (4, 2, 1),
    "U16_WB": (4, 2, 1),
    "I16": (3, 1, 1),
    "I16_B": (4, 1, 1),
    "I16_W": (4, 1, 1),
    "I16_WB": (4, 1, 1),
    "U32": (5, 2, 2),
    "U32_W": (6, 2, 2),
    "U32_B": (6, 2, 2),
    "U32_WB": (6, 2, 2),
    "I32": (5, 1, 2),
    "I32_W": (6, 1, 2),
    "I32_B": (6, 1, 2),
    "I32_WB": (6, 1, 2),
    "S32": (5, 1, 2),
    "S32_W": (6, 1, 2),
    "S32_B": (6, 1, 2),
    "S32_WB": (6, 1, 2),
    "F32": (7, 1, 2),
    "F32_W": (8, 1, 2),
    "F32_B": (8, 1, 2),
    "F32_WB": (8, 1, 2),
    "FLOAT32": (7, 1, 2),
    "FLOAT32_W": (8, 1, 2),
    "FLOAT32_WB": (8, 1, 2),
}

_VALID_SUN_ACTIONS = frozenset({"0", "1", "2", "4", "8"})


@dataclass
class _SunVar:
    orig_idx: str
    fc: int
    start_addr: int
    reg_count: int
    var_type: int
    var_signed: int
    bit_offset: int | None
    is_string: bool
    is_writable: bool
    name: str
    coeff_a: str
    coeff_b: str
    unit: str
    action: str


def _sanitize_ini_field(text: str) -> str:
    return text.replace("\r", " ").replace("\n", " ").replace(";", ",").strip()


def _map_action_to_webdynsun(raw_action: str) -> str:
    act = raw_action.strip()
    if act in _VALID_SUN_ACTIONS:
        return act
    if act in {"9", "10"}:
        return "8"  # Alarm on change of status
    if act in {"3", "5", "6", "7", "11", "12"}:
        return "4"  # Instantaneous measurement fallback
    return "4"


def _parse_addr_and_meta(
    raw_addr: str, dtype_u: str
) -> tuple[int, int, int, int, int | None, bool] | None:
    """Resolve (start_addr, reg_count, var_type, var_signed, bit_offset, is_string)."""
    parts = raw_addr.strip().split("_")
    try:
        start_addr = int(parts[0])
    except ValueError:
        return None
    if not (0 <= start_addr <= 65535):
        return None

    if dtype_u in {"STRING", "RAW"} or dtype_u.startswith("STR"):
        byte_len = 16
        if len(parts) > 1:
            try:
                byte_len = max(1, int(parts[1]))
            except ValueError:
                byte_len = 16
        elif dtype_u.startswith("STR") and dtype_u[3:].isdigit():
            byte_len = max(1, int(dtype_u[3:]))
        reg_count = max(1, (byte_len + 1) // 2)
        return start_addr, reg_count, 9, 2, None, True

    spec = _SUNPM_TO_SUN_TYPE.get(dtype_u)
    if spec is None:
        return None
    var_type, var_signed, reg_count = spec
    bit_offset: int | None = None
    if dtype_u == "BITS" and len(parts) >= 2:
        try:
            bit_offset = int(parts[1])
        except ValueError:
            bit_offset = 0
    return start_addr, reg_count, var_type, var_signed, bit_offset, False


def _format_coeff(raw_val: str, default: str) -> str:
    s = raw_val.strip()
    if not s:
        return default
    try:
        f = float(s)
        if not math.isfinite(f):
            return default
        if f == int(f):
            return str(int(f))
        return f"{f:g}"
    except ValueError:
        return default


def render_webdynsun_ini_from_rows(rows: list[list[str]]) -> tuple[str, list[str]]:
    """
    Build WebdynSun .ini content from WebdynSunPM 11-column data rows.

    Returns (ini_text, skipped_diagnostics).
    """
    parsed_vars: list[_SunVar] = []
    skipped: list[str] = []

    for row_num, cols in enumerate(rows, start=1):
        if len(cols) < 11:
            skipped.append(f"Row {row_num}: fewer than 11 columns")
            continue
        orig_idx = cols[0].strip() or str(row_num)
        fc_raw = cols[1].strip()
        try:
            fc = int(fc_raw)
        except ValueError:
            skipped.append(f"Row {row_num}: invalid function code '{fc_raw}'")
            continue
        if fc not in {1, 2, 3, 4}:
            skipped.append(f"Row {row_num}: unsupported function code '{fc}'")
            continue

        raw_addr = cols[2].strip()
        dtype_u = cols[3].strip().upper()
        name = _sanitize_ini_field(cols[5]) or f"Var_{row_num}"
        meta = _parse_addr_and_meta(raw_addr, dtype_u)
        if meta is None:
            skipped.append(
                f"Skipped unsupported WebdynSun type/address: {name} | addr={raw_addr} | type={dtype_u}"
            )
            continue

        start_addr, reg_count, var_type, var_signed, bit_offset, is_string = meta
        action = _map_action_to_webdynsun(cols[10])
        is_writable = action == "1" and fc in {1, 3}
        parsed_vars.append(
            _SunVar(
                orig_idx=orig_idx,
                fc=fc,
                start_addr=start_addr,
                reg_count=reg_count,
                var_type=var_type,
                var_signed=var_signed,
                bit_offset=bit_offset,
                is_string=is_string,
                is_writable=is_writable,
                name=name,
                coeff_a=_format_coeff(cols[7], "1"),
                coeff_b=_format_coeff(cols[8], "0"),
                unit=_sanitize_ini_field(cols[9]),
                action=action,
            )
        )

    requests_lines: list[str] = []
    variables_lines: list[str] = []

    req_idx = 0
    current_req: dict[str, Any] | None = None

    def _flush_req(req: dict[str, Any]) -> None:
        r_idx = req["index"]
        r_fc = req["fc"]
        r_start = req["start"]
        r_count = req["end"] - r_start
        r_write_fc = req["write_fc"]
        en_write = 1 if r_write_fc else 2
        requests_lines.append(
            f"{r_idx};Req{r_idx};{r_fc};{r_write_fc};{r_start};{r_count};1;{en_write};;"
        )
        for v_idx, v in req["vars"]:
            if v.fc in {1, 2} and v.var_type == 1:
                pos = (v.bit_offset if v.bit_offset is not None else 0) + 1
            else:
                pos = (v.start_addr - r_start) + 1
            variables_lines.append(
                f"{v_idx};{r_idx};{v.name};{v.var_type};{v.var_signed};{pos};;;"
                f"{v.coeff_a};{v.coeff_b};{v.unit};{v.action}"
            )

    for v_idx, item in enumerate(parsed_vars, start=1):
        item_end = item.start_addr + item.reg_count
        if item.is_writable:
            if current_req is not None:
                _flush_req(current_req)
                current_req = None
            req_idx += 1
            write_fc = 5 if item.fc == 1 else (6 if item.reg_count == 1 else 16)
            _flush_req(
                {
                    "index": req_idx,
                    "fc": item.fc,
                    "start": item.start_addr,
                    "end": item_end,
                    "write_fc": write_fc,
                    "vars": [(v_idx, item)],
                }
            )
            continue

        can_append = False
        if current_req is not None and not current_req["is_string"] and not item.is_string:
            if current_req["fc"] == item.fc:
                new_start = min(current_req["start"], item.start_addr)
                new_end = max(current_req["end"], item_end)
                # Pack contiguous or same-register slices (max MAX_WEBDYNSUN_READ_REGS)
                gap_ok = item.start_addr <= current_req["end"]
                if gap_ok and (new_end - new_start) <= MAX_WEBDYNSUN_READ_REGS:
                    can_append = True
                    current_req["start"] = new_start
                    current_req["end"] = new_end
                    current_req["vars"].append((v_idx, item))

        if not can_append:
            if current_req is not None:
                _flush_req(current_req)
            req_idx += 1
            current_req = {
                "index": req_idx,
                "fc": item.fc,
                "start": item.start_addr,
                "end": item_end,
                "write_fc": 0,
                "is_string": item.is_string,
                "vars": [(v_idx, item)],
            }

    if current_req is not None:
        _flush_req(current_req)

    header_notes = ""
    if skipped:
        note_lines = "\n".join(f"# NOTE: {s}" for s in skipped[:25])
        if len(skipped) > 25:
            note_lines += f"\n# NOTE: ... and {len(skipped) - 25} more skipped rows"
        header_notes = note_lines + "\n"

    content = (
        f"{header_notes}"
        "Modbus_RequestsTables={\n"
        "# reqIndex;reqName;reqReadFctCode;reqWriteFctCode;reqStartRegister;reqNbRegisters;reqEnableReading;reqEnableWriting;reqOption1;reqOption2\n"
        + ("\n".join(requests_lines) + "\n" if requests_lines else "")
        + "}\n\n"
        "Modbus_VariablesTables={\n"
        "# varIndex;varReqIndex;varName;varType;varSigned;varPosition;varOption1;varOption2;coeffA;coeffB;unit;varAction\n"
        + ("\n".join(variables_lines) + "\n" if variables_lines else "")
        + "}\n"
    )
    return content.replace("\r\n", "\n").replace("\n", "\r\n"), skipped


def convert_sunpm_csv_to_webdynsun_ini(
    csv_path: Path, ini_path: Path
) -> tuple[bool, list[str]]:
    """Convert a WebdynSunPM 11-column CSV file to a WebdynSun .ini file atomically."""
    with open(csv_path, encoding="utf-8-sig", newline="") as fp:
        all_rows = list(csv.reader(fp, delimiter=";"))
    data_rows = all_rows[1:] if len(all_rows) > 1 else []
    content, skipped = render_webdynsun_ini_from_rows(data_rows)
    with staged_text_output(str(ini_path), encoding="utf-8") as fp:
        fp.write(content)
    return True, skipped


def verify_webdynsun_ini(ini_path: Path) -> tuple[bool, list[str]]:
    """Validate structure and field invariants of a WebdynSun .ini definition file."""
    issues: list[str] = []
    try:
        raw = ini_path.read_bytes()
    except OSError as exc:
        return False, [f"Cannot read {ini_path.name}: {exc}"]
    if raw.startswith(b"\xef\xbb\xbf"):
        issues.append(f"{ini_path.name}: contains prohibited UTF-8 BOM")
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        return False, [f"{ini_path.name}: invalid UTF-8 ({exc})"]

    if "Modbus_RequestsTables={" not in text or "Modbus_VariablesTables={" not in text:
        issues.append(f"{ini_path.name}: missing Modbus_RequestsTables or Modbus_VariablesTables")
        return False, issues

    req_indices: set[int] = set()
    in_req = False
    in_var = False

    for line_no, line in enumerate(text.splitlines(), start=1):
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        if s.startswith("Modbus_RequestsTables="):
            in_req, in_var = True, False
            continue
        if s.startswith("Modbus_VariablesTables="):
            in_req, in_var = False, True
            continue
        if s == "}":
            in_req = in_var = False
            continue

        cols = s.split(";")
        if in_req:
            if len(cols) < 8:
                issues.append(f"{ini_path.name}:{line_no}: request row < 8 fields")
                continue
            try:
                r_idx = int(cols[0])
                r_fc = int(cols[2])
                r_start = int(cols[4])
                r_nb = int(cols[5])
            except ValueError:
                issues.append(f"{ini_path.name}:{line_no}: non-integer request field")
                continue
            req_indices.add(r_idx)
            if r_fc not in {0, 1, 2, 3, 4}:
                issues.append(f"{ini_path.name}:{line_no}: invalid read FC {r_fc}")
            if not (0 <= r_start <= 65535):
                issues.append(f"{ini_path.name}:{line_no}: startRegister {r_start} out of bounds")
            if not (1 <= r_nb <= 125):
                issues.append(f"{ini_path.name}:{line_no}: reqNbRegisters {r_nb} out of 1..125")
        elif in_var:
            if len(cols) < 12:
                issues.append(f"{ini_path.name}:{line_no}: variable row < 12 fields")
                continue
            try:
                v_req = int(cols[1])
                v_type = int(cols[3])
                v_sign = int(cols[4])
                v_pos = int(cols[5])
            except ValueError:
                issues.append(f"{ini_path.name}:{line_no}: non-integer variable field")
                continue
            if v_req not in req_indices:
                issues.append(f"{ini_path.name}:{line_no}: unknown varReqIndex {v_req}")
            if v_type not in range(1, 11):
                issues.append(f"{ini_path.name}:{line_no}: invalid varType {v_type}")
            if v_sign not in {1, 2}:
                issues.append(f"{ini_path.name}:{line_no}: invalid varSigned {v_sign}")
            if v_pos < 1:
                issues.append(f"{ini_path.name}:{line_no}: invalid varPosition {v_pos}")
            if cols[11].strip() not in _VALID_SUN_ACTIONS:
                issues.append(f"{ini_path.name}:{line_no}: invalid varAction '{cols[11]}'")

    return len(issues) == 0, issues
