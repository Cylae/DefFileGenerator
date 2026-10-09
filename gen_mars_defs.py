#!/usr/bin/env python3
"""Generate WebdynSunPM Modbus TCP definition files for the MarsRenewable Local Controller (LC).

Source document : MarsLocalController_IO_List_ModbusTCP_V7_0.xlsx
Output files    : one definition file per Modbus slave ID, because a WebdynSunPM
                  definition file describes exactly one slave.

    MarsRenewable_LC_V7.csv      slave 1      PCS, bank, meter, AC, DIO, PV, liquid cooling, alarm summary
    MarsRenewable_LC_EMS_V7.csv  slave 247    LC overview (FC4) + EMS write block (FC3 / forced write FC16)
    MarsRenewable_Rack_V7.csv    slave N+1    one rack + its cells (declare one device per rack)

File format (reference: HUAWEI_V4.csv, validated with DefFileGenerator):
    line 1 : Protocol;Category;Manufacturer;Model;ForcedWriteCode;;;;;;
    rows   : Index;Info1(FC);Info2(address);Info3(type);Info4;Name;Tag;CoefA;CoefB;Unit;Action
    UTF-8 without BOM, CRLF line endings, ';' separator.

Action codes used:
    0 disabled | 1 parameter (read once) | 4 instant value | 9 instant value + alarm on change

Scaling: the IO list states that no scaling is required, so CoefA = 1 / CoefB = 0 everywhere
except where the IO list gives a different unit (0.01 Hz) or a standard WebdynSunPM tag needs W / Var.

Scaling the number of equipment instances is done on the command line, e.g.:
    python gen_mars_defs.py --pcs 4 --bank 4 --meter 2 --cells 52
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

FC_DISCRETE, FC_HOLDING, FC_INPUT = 2, 3, 4
ACT_OFF, ACT_PARAM, ACT_INST, ACT_ALARM = 0, 1, 4, 9

PROTOCOL = "modbusTCP"
CATEGORY = "Inverter"
MANUFACTURER = "MarsRenewable"
VERSION = "V7"

# Register layout of the IO list: (first address, stride between two instances, max instances)
LAYOUT = {
    "pcs": (0, 100, 50),
    "bank": (5000, 100, 50),
    "meter": (10000, 50, 100),
    "ac": (15000, 50, 100),
    "dio": (20000, 22, 50),  # stride deduced from the 22-register 1#DIO block
    "pv": (25000, 100, 50),
    "cooling": (30000, 40, 100),
}
BRANCH_FIRST_OFFSET, BRANCH_STRIDE, BRANCH_MAX = 62, 6, 6  # 62 + 6 * 6 = 98 <= 100 registers
CELL_FIRST_ADDR, CELL_STRIDE, CELL_MAX = 100, 10, 1000


@dataclass(frozen=True)
class Row:
    fc: int
    addr: int
    typ: str
    name: str
    tag: str
    unit: str = ""
    coef_a: float = 1.0
    coef_b: float = 0.0
    action: int = ACT_INST


def _num(value: float) -> str:
    return f"{value:g}"


def write_definition(path: Path, model: str, rows: list[Row], forced_write: str = "") -> None:
    tags = [r.tag for r in rows]
    dupes = {t for t in tags if tags.count(t) > 1}
    if dupes:
        raise ValueError(f"duplicate tags in {path.name}: {sorted(dupes)}")
    lines = [
        ";".join([PROTOCOL, CATEGORY, MANUFACTURER, model, forced_write, "", "", "", "", "", ""])
    ]
    for index, r in enumerate(rows, start=1):
        for text in (r.name, r.tag, r.unit):
            if ";" in text or "\n" in text:
                raise ValueError(f"illegal character in field {text!r}")
        lines.append(
            ";".join(
                [
                    str(index),
                    str(r.fc),
                    str(r.addr),
                    r.typ,
                    "",
                    r.name,
                    r.tag,
                    _num(r.coef_a),
                    _num(r.coef_b),
                    r.unit,
                    str(r.action),
                ]
            )
        )
    path.write_bytes(("\r\n".join(lines) + "\r\n").encode("utf-8"))


# --------------------------------------------------------------------------- slave 1 blocks
def alarm_summary_rows() -> list[Row]:
    names = ["LC", "PCS", "Bank", "Air Conditioning"]
    return [
        Row(
            FC_DISCRETE,
            i,
            "U16",
            f"{n} Alarm Summary [0=normal, 1=alarm]",
            f"alarm_{n.lower().replace(' ', '_')}",
            action=ACT_ALARM,
        )
        for i, n in enumerate(names)
    ]


def ac_power_rows(
    base: int, label: str, tag: str, branches: int, std_tags: bool, line_voltage: bool
) -> list[Row]:
    """PCS and PV blocks share the same register layout (offsets 2..61 + DC branches)."""
    t = tag
    r: list[Row] = []

    def f32(off: int, name: str, key: str, unit: str, action: int = ACT_INST, **kw) -> None:
        r.append(
            Row(
                FC_INPUT,
                base + off,
                "F32",
                f"{label} {name}",
                f"{t}_{key}",
                unit,
                action=action,
                **kw,
            )
        )

    stat_tag = "status" if std_tags else f"{t}_status"
    r.append(
        Row(
            FC_INPUT,
            base,
            "U16",
            f"{label} Operating Status [0=stopped, 1=running]",
            stat_tag,
            action=ACT_ALARM,
        )
    )
    if line_voltage:  # PCS sheet: line-to-line voltages
        volts = (
            (2, "AC Uab Line Voltage", "uab"),
            (4, "AC Ubc Line Voltage", "ubc"),
            (6, "AC Uca Line Voltage", "uca"),
        )
    else:  # PV sheet: phase voltages
        volts = (
            (2, "AC Phase A Voltage", "ua"),
            (4, "AC Phase B Voltage", "ub"),
            (6, "AC Phase C Voltage", "uc"),
        )
    for off, name, key in volts:
        f32(off, name, key, "V")
    for off, name, key in (
        (8, "AC Phase A Current", "ia"),
        (10, "AC Phase B Current", "ib"),
        (12, "AC Phase C Current", "ic"),
    ):
        f32(off, name, key, "A")
    f32(14, "AC Frequency", "freq", "Hz")
    if std_tags:
        r.append(
            Row(
                FC_INPUT,
                base + 16,
                "F32",
                f"{label} Total AC Active Power",
                "RealPower",
                "W",
                coef_a=1000,
            )
        )
        r.append(
            Row(
                FC_INPUT,
                base + 18,
                "F32",
                f"{label} Total AC Reactive Power",
                "ReactivePower",
                "Var",
                coef_a=1000,
            )
        )
    else:
        f32(16, "Total AC Active Power [+charge, -discharge]", "p_total", "kW")
        f32(18, "Total AC Reactive Power", "q_total", "kvar")
    f32(20, "Total AC Apparent Power", "s_total", "kVA")
    f32(22, "Total AC Power Factor", "pf_total", "")
    for k, ph in enumerate("ABC"):
        f32(24 + 2 * k, f"Phase {ph} AC Active Power", f"p_{ph.lower()}", "kW")
    for k, ph in enumerate("ABC"):
        f32(30 + 2 * k, f"Phase {ph} AC Reactive Power", f"q_{ph.lower()}", "kvar")
    for k, ph in enumerate("ABC"):
        f32(36 + 2 * k, f"Phase {ph} AC Apparent Power", f"s_{ph.lower()}", "kVA")
    for k, ph in enumerate("ABC"):
        f32(42 + 2 * k, f"Phase {ph} AC Power Factor", f"pf_{ph.lower()}", "")
    f32(48, "IGBT Module Temperature", "igbt_temp", "°C")
    f32(50, "Ambient Temperature", "amb_temp", "°C")
    f32(52, "Total AC Charging Energy", "e_chg_total", "kWh")
    f32(54, "Total AC Discharging Energy", "e_dis_total", "kWh")
    f32(56, "Daily AC Charging Energy", "e_chg_day", "kWh")
    f32(58, "Daily AC Discharging Energy", "e_dis_day", "kWh")
    r.append(
        Row(
            FC_INPUT,
            base + 60,
            "U32",
            f"{label} Number of branches",
            f"{t}_branches",
            action=ACT_PARAM,
        )
    )
    for b in range(1, branches + 1):
        off = BRANCH_FIRST_OFFSET + (b - 1) * BRANCH_STRIDE
        f32(off, f"Branch {b} DC Power", f"br{b}_p", "kW")
        f32(off + 2, f"Branch {b} DC Voltage", f"br{b}_v", "V")
        f32(off + 4, f"Branch {b} DC Current", f"br{b}_i", "A")
    return r


def bank_rows(base: int, label: str, t: str) -> list[Row]:
    spec = [
        (0, "Bank Voltage", "v", "V", ACT_INST),
        (2, "Bank Current", "i", "A", ACT_INST),
        (4, "Bank SOC", "soc", "%", ACT_INST),
        (6, "Bank SOH", "soh", "%", ACT_INST),
        (8, "Bank Insulation Resistance", "insulation", "Ohm", ACT_INST),
        (10, "Bank Available Charge Energy", "e_avail_chg", "kWh", ACT_INST),
        (12, "Bank Available Discharge Energy", "e_avail_dis", "kWh", ACT_INST),
        (14, "System Allowed Maximum Charge Current", "i_max_chg", "A", ACT_INST),
        (16, "System Allowed Maximum Discharge Current", "i_max_dis", "A", ACT_INST),
        (18, "Today's Charge Energy", "e_chg_day", "kWh", ACT_INST),
        (20, "Today's Discharge Energy", "e_dis_total", "kWh", ACT_INST),
        (22, "Accumulated Charge Energy", "e_chg_accum", "kWh", ACT_INST),
        (24, "Accumulated Discharge Energy", "e_dis_accum", "kWh", ACT_INST),
    ]
    return [
        Row(FC_INPUT, base + o, "F32", f"{label} {n}", f"{t}_{k}", u, action=a)
        for o, n, k, u, a in spec
    ]


def meter_rows(base: int, label: str, t: str) -> list[Row]:
    tou = ["Active", "Sharp", "Peak", "Flat", "Valley"]
    rows: list[Row] = []
    for d, direction in ((0, "Forward"), (10, "Reverse")):
        for k, kind in enumerate(tou):
            rows.append(
                Row(
                    FC_INPUT,
                    base + d + 2 * k,
                    "F32",
                    f"{label} Total {direction} {kind} Energy",
                    f"{t}_{direction.lower()[:3]}_{kind.lower()}",
                    "kWh",
                )
            )
    rows.append(
        Row(FC_INPUT, base + 20, "F32", f"{label} Total Active Power", f"{t}_p_total", "kW")
    )
    return rows


def ac_unit_rows(base: int, label: str, t: str) -> list[Row]:
    rows = [
        Row(
            FC_INPUT,
            base,
            "U16",
            f"{label} Operating Status [0=stopped, 1=running]",
            f"{t}_status",
            action=ACT_ALARM,
        )
    ]
    setpoints = [
        (4, "Cooling Low Point", "cool_low", "°C"),
        (6, "Cooling High Point", "cool_high", "°C"),
        (8, "Heating Low Point", "heat_low", "°C"),
        (10, "Heating High Point", "heat_high", "°C"),
        (12, "Dehumidification High Point (Start Point)", "dehum_start", "%"),
        (14, "Dehumidification Low Point (Stop Point)", "dehum_stop", "%"),
        (16, "High Temperature Alarm Point", "t_alarm_high", "°C"),
        (18, "Low Temperature Alarm Point", "t_alarm_low", "°C"),
    ]
    rows += [
        Row(FC_INPUT, base + o, "F32", f"{label} {n}", f"{t}_{k}", u, action=ACT_PARAM)
        for o, n, k, u in setpoints
    ]
    rows.append(Row(FC_INPUT, base + 20, "F32", f"{label} Temperature", f"{t}_temp", "°C"))
    flags = [
        (22, "Temperature Status", "st_temp"),
        (23, "Voltage Status", "st_volt"),
        (24, "Sensor Status", "st_sensor"),
        (25, "Compressor Status", "st_comp"),
        (26, "Air Pressure Status", "st_press"),
    ]
    rows += [
        Row(
            FC_INPUT,
            base + o,
            "U16",
            f"{label} {n} [0=normal, 1=alarm]",
            f"{t}_{k}",
            action=ACT_ALARM,
        )
        for o, n, k in flags
    ]
    rows.append(Row(FC_INPUT, base + 27, "F32", f"{label} Relative Humidity", f"{t}_rh", "%"))
    return rows


def dio_rows(base: int, label: str, t: str) -> list[Row]:
    di = {
        1: ("crash-stop", 0),
        2: ("Surge fault feedback", 0),
        3: ("PCS Discharge/Charge Status", 0),
        4: ("Water Cooling", 0),
        5: ("Aerosol spraying", 0),
        6: ("KA1-KA6 Contactor Monitoring", 0),
        9: ("Temperature feedback", 1),
        10: ("Smoke feedback", 1),
        11: ("AC side entrance guard", 1),
        12: ("DC side entrance guard", 1),
        13: ("Exhaust Fan Feedback", 1),
    }  # DI-7 / DI-8 are reserved; value = 1 when "1" means normal (inverted logic)
    rows: list[Row] = []
    for n, (name, inverted) in di.items():
        logic = "1=normal, 0=alarm" if inverted else "0=normal, 1=alarm"
        rows.append(
            Row(
                FC_INPUT,
                base + 5 + n,
                "U16",
                f"{label} DI-{n} ({name}) [{logic}]",
                f"{t}_di{n}",
                action=ACT_ALARM,
            )
        )
    do = {1: "Running Indicator", 2: "Fault Indicator", 3: "Intermediate relay"}
    for n, name in do.items():
        rows.append(
            Row(
                FC_INPUT,
                base + 18 + n,
                "U16",
                f"{label} DO-{n} ({name}) [0=off, 1=on]",
                f"{t}_do{n}",
            )
        )
    return rows


def cooling_rows(base: int, label: str, t: str) -> list[Row]:
    rows = [
        Row(
            FC_INPUT,
            base,
            "U16",
            f"{label} Operating Status [0=stopped, 1=running]",
            f"{t}_status",
            action=ACT_ALARM,
        ),
        Row(
            FC_INPUT,
            base + 4,
            "U16",
            f"{label} Operating Mode [0=stop,1=cool,2=heat,3=self-circ,4=auto,5=standby]",
            f"{t}_mode",
        ),
        Row(FC_INPUT, base + 6, "F32", f"{label} Ambient Temperature", f"{t}_amb_temp", "°C"),
        Row(
            FC_INPUT, base + 8, "F32", f"{label} Supply Water Temperature", f"{t}_supply_temp", "°C"
        ),
        Row(
            FC_INPUT, base + 10, "F32", f"{label} Inlet Water Temperature", f"{t}_inlet_temp", "°C"
        ),
        Row(FC_INPUT, base + 12, "F32", f"{label} Supply Water Pressure", f"{t}_supply_press"),
        Row(FC_INPUT, base + 14, "F32", f"{label} Return Water Pressure", f"{t}_return_press"),
        Row(
            FC_INPUT,
            base + 16,
            "U16",
            f"{label} Fan Operating Status",
            f"{t}_fan_status",
            action=ACT_ALARM,
        ),
        Row(
            FC_INPUT,
            base + 17,
            "U16",
            f"{label} Water Pump Operating Status",
            f"{t}_pump_status",
            action=ACT_ALARM,
        ),
        Row(
            FC_INPUT,
            base + 18,
            "U16",
            f"{label} Electric Heating Operating Status",
            f"{t}_heater_status",
            action=ACT_ALARM,
        ),
        Row(
            FC_INPUT,
            base + 19,
            "U16",
            f"{label} Refrigerant Operating Status",
            f"{t}_refrig_status",
            action=ACT_ALARM,
        ),
        Row(FC_INPUT, base + 20, "F32", f"{label} Fan Output", f"{t}_fan_output"),
        Row(FC_INPUT, base + 22, "F32", f"{label} Water Pump Output", f"{t}_pump_output"),
    ]
    for off, name, key in (
        (24, "Target Cooling Start Point", "cool_start"),
        (26, "Target Cooling Stop Point", "cool_stop"),
        (28, "Target Heating Start Point", "heat_start"),
        (30, "Target Heating Stop Point", "heat_stop"),
    ):
        rows.append(
            Row(
                FC_INPUT, base + off, "F32", f"{label} {name}", f"{t}_{key}", "°C", action=ACT_PARAM
            )
        )
    rows.append(
        Row(
            FC_INPUT,
            base + 32,
            "U16",
            f"{label} EMS Control Mode Status [0=false, 1=true]",
            f"{t}_ems_ctrl",
        )
    )
    return rows


def build_slave1(n: dict[str, int], branches: int) -> list[Row]:
    def base(kind: str, i: int) -> int:
        first, stride, _ = LAYOUT[kind]
        return first + (i - 1) * stride

    rows = alarm_summary_rows()
    for i in range(1, n["pcs"] + 1):
        rows += ac_power_rows(
            base("pcs", i), f"PCS{i}", f"pcs{i}", branches, std_tags=False, line_voltage=True
        )
    for i in range(1, n["bank"] + 1):
        rows += bank_rows(base("bank", i), f"Bank{i}", f"bank{i}")
    for i in range(1, n["meter"] + 1):
        rows += meter_rows(base("meter", i), f"Meter{i}", f"meter{i}")
    for i in range(1, n["ac"] + 1):
        rows += ac_unit_rows(base("ac", i), f"AC{i}", f"ac{i}")
    for i in range(1, n["dio"] + 1):
        rows += dio_rows(base("dio", i), f"DIO{i}", f"dio{i}")
    for i in range(1, n["pv"] + 1):
        # standard WebdynSunPM tags (status / RealPower / ReactivePower) only on PV1
        rows += ac_power_rows(
            base("pv", i), f"PV{i}", f"pv{i}", branches, std_tags=(i == 1), line_voltage=False
        )
    for i in range(1, n["cooling"] + 1):
        rows += cooling_rows(base("cooling", i), f"LiquidCooling{i}", f"lcool{i}")
    return rows


# --------------------------------------------------------------------------- slave 247
def build_ems() -> list[Row]:
    rows = [
        Row(
            FC_INPUT,
            0,
            "U16",
            "LC Measurement Point Version Number",
            "lc_map_version",
            action=ACT_PARAM,
        ),
        Row(FC_INPUT, 1, "U16", "LC PCS Number", "lc_pcs_count", action=ACT_PARAM),
        Row(FC_INPUT, 2, "U16", "LC Bank Number", "lc_bank_count", action=ACT_PARAM),
        Row(FC_INPUT, 3, "U16", "LC Rack Number", "lc_rack_count", action=ACT_PARAM),
        Row(FC_INPUT, 4, "U16", "LC Meter Number", "lc_meter_count", action=ACT_PARAM),
        Row(FC_HOLDING, 0, "U16", "EMS Working Mode [0=autonomous, 1=remote control]", "ems_mode"),
        Row(FC_HOLDING, 1, "I32", "Remote Control Duration", "ems_duration", "min"),
        Row(
            FC_HOLDING,
            3,
            "U16",
            "Energy Storage Working Mode [0=grid constant power, 1=off-grid constant voltage, 2=stop]",
            "ess_mode",
        ),
        Row(
            FC_HOLDING,
            4,
            "I32",
            "Energy Storage Active Power [<0 discharge, >0 charge]",
            "ess_p_set",
            "W",
        ),
        Row(
            FC_HOLDING,
            6,
            "I32",
            "Energy Storage Reactive Power [not supported]",
            "ess_q_set",
            "Var",
            action=ACT_OFF,
        ),
        Row(FC_HOLDING, 8, "U16", "AC Voltage [off-grid mode]", "ess_v_set", "V"),
        Row(FC_HOLDING, 9, "U16", "AC Frequency [off-grid mode]", "ess_f_set", "Hz", coef_a=0.01),
    ]
    return rows


# --------------------------------------------------------------------------- rack (slave N+1)
def build_rack(cells: int) -> list[Row]:
    rows = [Row(FC_INPUT, 0, "U16", "Number of Cells in Rack", "rack_cells", action=ACT_PARAM)]
    spec = [
        (1, "Rack Voltage", "v", "V"),
        (3, "Rack Current", "i", "A"),
        (5, "Rack SOC", "soc", "%"),
        (7, "Rack SOH", "soh", "%"),
        (9, "Rack Insulation Resistance", "insulation", "Ohm"),
        (11, "Maximum Allowable Charge Current", "i_max_chg", "A"),
        (13, "Maximum Allowable Discharge Current", "i_max_dis", "A"),
        (15, "Total Charge Energy", "e_chg_total", "kWh"),
        (17, "Total Discharge Energy", "e_dis_total", "kWh"),
        (19, "Available Charge Energy", "e_avail_chg", "kWh"),
        (21, "Available Discharge Energy", "e_avail_dis", "kWh"),
        (23, "Average Cell Voltage", "cell_v_avg", "V"),
        (25, "Average Cell Temperature", "cell_t_avg", "°C"),
    ]
    for off, name, key, unit in spec:
        rows.append(Row(FC_INPUT, off, "F32", f"Rack {name}", f"rack_{key}", unit))
    extremes = [
        (27, "Cell Voltage", "cell_v", "V"),
        (35, "Cell Temperature", "cell_t", "°C"),
        (43, "Cell SOC", "cell_soc", "%"),
        (51, "Cell SOH", "cell_soh", "%"),
    ]
    for off, name, key, unit in extremes:
        for k, (kind, short) in enumerate((("Maximum", "max"), ("Minimum", "min"))):
            a = off + 4 * k
            rows.append(Row(FC_INPUT, a, "F32", f"Rack {kind} {name}", f"rack_{key}_{short}", unit))
            rows.append(
                Row(
                    FC_INPUT,
                    a + 2,
                    "F32",
                    f"Rack {kind} {name} Cell Number",
                    f"rack_{key}_{short}_no",
                )
            )
    for c in range(1, cells + 1):
        base = CELL_FIRST_ADDR + (c - 1) * CELL_STRIDE
        rows += [
            Row(FC_INPUT, base, "F32", f"Cell {c} Voltage", f"cell{c}_v", "V"),
            Row(FC_INPUT, base + 2, "F32", f"Cell {c} Temperature", f"cell{c}_t", "°C"),
            Row(FC_INPUT, base + 4, "F32", f"Cell {c} SOC", f"cell{c}_soc", "%"),
            Row(FC_INPUT, base + 6, "F32", f"Cell {c} SOH", f"cell{c}_soh", "%"),
            Row(FC_INPUT, base + 8, "F32", f"Cell {c} Resistance", f"cell{c}_r", "Ohm"),
        ]
    return rows


def main() -> None:
    p = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    for kind, (_, _, maximum) in LAYOUT.items():
        p.add_argument(
            f"--{kind}", type=int, default=1, help=f"instances of {kind} (0..{maximum}, default 1)"
        )
    p.add_argument(
        "--branches", type=int, default=BRANCH_MAX, help=f"DC branches per PCS/PV (1..{BRANCH_MAX})"
    )
    p.add_argument(
        "--cells", type=int, default=16, help=f"cells per rack (0..{CELL_MAX}, default 16)"
    )
    p.add_argument("--out-dir", type=Path, default=Path("."))
    a = p.parse_args()

    counts = {k: getattr(a, k) for k in LAYOUT}
    for k, v in counts.items():
        if not 0 <= v <= LAYOUT[k][2]:
            p.error(f"--{k} must be within 0..{LAYOUT[k][2]}")
    if not 1 <= a.branches <= BRANCH_MAX or not 0 <= a.cells <= CELL_MAX:
        p.error("--branches / --cells out of range")

    a.out_dir.mkdir(parents=True, exist_ok=True)
    jobs = [
        (f"MarsRenewable_LC_{VERSION}.csv", f"LC_{VERSION}", build_slave1(counts, a.branches), ""),
        (f"MarsRenewable_LC_EMS_{VERSION}.csv", f"LC_EMS_{VERSION}", build_ems(), "16"),
        (f"MarsRenewable_Rack_{VERSION}.csv", f"Rack_{VERSION}", build_rack(a.cells), ""),
    ]
    for filename, model, rows, forced in jobs:
        write_definition(a.out_dir / filename, model, rows, forced)
        print(f"{filename}: {len(rows)} variables")


if __name__ == "__main__":
    main()
