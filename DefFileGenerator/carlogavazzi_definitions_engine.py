"""
Apollo-Standard Master Definition Engine for Carlo Gavazzi Energy Meters & Power Analyzers.

Official Documentation Sources:
  1. EM210_MV: EM210_MV - CP Ver.1 Rev.4 - 20220826.pdf
  2. EM210: em210-modbus-table.pdf & CarloGavazzi_EM210.ini
  3. EM530 & EM540: EM530-EM540 COMMUNICATION PROTOCOL Rev. 1.3 (June 03rd, 2021) & CarloGavazzi_EM530_test.csv
  4. EM111, EM112, ET112: EM100 Series Communication Protocol (em111_em112_et112_cp.pdf)
  5. EM330, EM340, ET330, ET340: EM300 Series Communication Protocol (em330_em340_et330_et340_cp.pdf)
  6. EM21: EM21 COMMUNICATION PROTOCOL V1r2 2010 (EM21 COMMUNICATION PROTOCOL_V1r2 2010.pdf)
  7. EM271: EM271 CP V0R0 ENG (EM271 CP V0R0 ENG.PDF)
  8. WM15: WM15-modbus_registers-V1-1-20192510.pdf
  9. WM20: WM20-modbus_registers.pdf
  10. WM50: wm50_modbus_registers.pdf

Strictly enforces:
  - Pillar 1: Strict UTF-8 without BOM.
  - Pillar 2: Strict CRLF (\\r\\n) line endings.
  - Pillar 3: Exact 11-column semicolon CSV schema with standard Webdyn header.
  - Pillar 4: Monotonic sequential 1..N row indexing.
  - Pillar 5: Valid Webdyn data types (I16, U16, I32, I32_W, U32, I64, U64, FLOAT)
             and unique canonical SunPM tags (P_AC, Q_AC, S_AC, FREQ, PF, E_TOT, etc.).
"""

from __future__ import annotations

from collections.abc import Sequence
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


class CarloGavazziDefinitionsEngine:
    """Master engine for compiling certified Carlo Gavazzi energy meter definitions."""

    MANUFACTURER: str = "CarloGavazzi"
    DEVICE_TYPE: str = "Meter"
    PROTOCOL: str = "Modbus"

    @classmethod
    def format_row(cls, index: int, reg: ModbusRegisterDef) -> str:
        """Format a single register definition into an 11-column semicolon row."""
        return (
            f"{index};{reg.fc};{reg.address};{reg.data_type};{reg.reserved};"
            f"{reg.name};{reg.tag};{reg.coef_a};{reg.coef_b};{reg.unit};{reg.action}"
        )

    @classmethod
    def format_header(cls, model: str) -> str:
        """Format the 11-column header for a Webdyn Modbus definition file."""
        return f"{cls.PROTOCOL};{cls.DEVICE_TYPE};{cls.MANUFACTURER};{model};;;;;;;"

    @classmethod
    def write_definition_file(
        cls,
        output_path: Path,
        model: str,
        registers: Sequence[ModbusRegisterDef],
    ) -> None:
        """Write a definition file strictly adhering to the 5 Webdyn Pillars."""
        output_path.parent.mkdir(parents=True, exist_ok=True)
        lines = [cls.format_header(model)]
        for idx, reg in enumerate(registers, 1):
            lines.append(cls.format_row(idx, reg))

        content = "\r\n".join(lines) + "\r\n"
        output_path.write_bytes(content.encode("utf-8"))

    # =========================================================================
    # 1. EM210 & EM210_MV (1-based Modicom addressing, exact 115 registers)
    # =========================================================================

    @classmethod
    def get_em210_registers(cls, is_mv: bool = False) -> list[ModbusRegisterDef]:
        """
        Build register definitions for Carlo Gavazzi EM210 and EM210_MV.
        Sources:
          - EM210_MV - CP Ver.1 Rev.4 - 20220826.pdf
          - em210-modbus-table.pdf & CarloGavazzi_EM210.ini
        """
        regs: list[ModbusRegisterDef] = []

        # Section 2.3 Instantaneous variables and meters (grouped by variable type)
        # Voltages L-N
        regs.append(
            ModbusRegisterDef(
                "3", "1", "I32", "", "volt1", "U_AC_1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "3", "I32", "", "volt2", "U_AC_2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "5", "I32", "", "volt3", "U_AC_3", "0.100000", "0.000000", "V", "4"
            )
        )
        # Currents
        regs.append(
            ModbusRegisterDef(
                "3", "13", "I32", "", "current1", "I_AC_1", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "15", "I32", "", "current2", "I_AC_2", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "17", "I32", "", "current3", "I_AC_3", "0.001000", "0.000000", "A", "4"
            )
        )
        # Active Powers
        regs.append(
            ModbusRegisterDef(
                "3", "19", "I32", "", "power1", "P_AC_1", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "21", "I32", "", "power2", "P_AC_2", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "23", "I32", "", "power3", "P_AC_3", "0.100000", "0.000000", "W", "4"
            )
        )
        # Apparent Powers
        regs.append(
            ModbusRegisterDef(
                "3", "25", "I32", "", "apparent_power1", "S_AC_1", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "27", "I32", "", "apparent_power2", "S_AC_2", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "29", "I32", "", "apparent_power3", "S_AC_3", "0.100000", "0.000000", "VA", "4"
            )
        )
        # Reactive Powers
        regs.append(
            ModbusRegisterDef(
                "3",
                "31",
                "I32",
                "",
                "reactive_power1",
                "Q_AC_1",
                "0.100000",
                "0.000000",
                "VAR",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "33",
                "I32",
                "",
                "reactive_power2",
                "Q_AC_2",
                "0.100000",
                "0.000000",
                "VAR",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "35",
                "I32",
                "",
                "reactive_power3",
                "Q_AC_3",
                "0.100000",
                "0.000000",
                "VAR",
                "4",
            )
        )
        # System Totals
        regs.append(
            ModbusRegisterDef(
                "3", "37", "I32", "", "volt", "U_AC", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "39", "I32", "", "V l-l", "U_AC_L1L2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "41", "I32", "", "power", "P_AC", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "43", "I32", "", "apparent_power", "S_AC", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "45", "I32", "", "reactive_power", "Q_AC", "0.100000", "0.000000", "VAR", "4"
            )
        )
        # Power Factors
        regs.append(
            ModbusRegisterDef(
                "3", "47", "I16", "", "Pf l1", "PF_1", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "48", "I16", "", "Pf l2", "PF_2", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "49", "I16", "", "Pf l3", "PF_3", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "50", "I16", "", "Pf", "PF", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "51", "I16", "", "Phase sequence", "", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "52", "I16", "", "frequency", "FREQ", "1.000000", "0.000000", "Hz", "4"
            )
        )
        # Energies
        regs.append(
            ModbusRegisterDef(
                "3", "53", "I32_W", "", "energy_tot", "E_TOT", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "55",
                "I32",
                "",
                "reactive_energy_tot",
                "Q_TOT",
                "0.100000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "79", "I32", "", "kWh(-) TOT", "E_EXP", "0.100000", "0.000000", "kWh", "4"
            )
        )
        # Operating hours
        regs.append(
            ModbusRegisterDef(
                "3", "91", "I32", "", "Hour counter meter", "", "0.010000", "0.000000", "Hours", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "93",
                "I32",
                "",
                "Hour counter meter neg",
                "",
                "0.010000",
                "0.000000",
                "Hours",
                "4",
            )
        )
        # Harmonic Distortions (THD)
        regs.append(
            ModbusRegisterDef(
                "3", "131", "I32", "", "THD(15h) A L1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "133", "I32", "", "THD(15h) A L2", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "135", "I32", "", "THD(15h) A L3", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "139", "I32", "", "THD(15h) V1-N", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "141", "I32", "", "THD(15h) V2-N", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "143", "I32", "", "THD(15h) V3-N", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "147", "I32", "", "THD(15h) V12", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "149", "I32", "", "THD(15h) V23", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "151", "I32", "", "THD(15h) V31", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "153",
                "I32",
                "",
                "Neutral current (In)",
                "I_AC_N",
                "0.001000",
                "0.000000",
                "A",
                "4",
            )
        )

        # Section 2.4 Duplicate registers (grouped by phase, untagged for tag uniqueness)
        regs.append(
            ModbusRegisterDef(
                "3",
                "247",
                "I32",
                "",
                "Hour counter meter neg_1",
                "",
                "0.010000",
                "0.000000",
                "Hours",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "249",
                "I32",
                "",
                "Neutral current (In)_1",
                "",
                "0.001000",
                "0.000000",
                "A",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "255",
                "I32",
                "",
                "Hour counter meter_1",
                "",
                "0.010000",
                "0.000000",
                "Hours",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "259", "I32", "", "V L-N sum", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "261", "I32", "", "V L-L sum", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "263", "I32", "", "W sum", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "265", "I32", "", "VA sum", "", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "267", "I32", "", "VAR sum", "", "0.100000", "0.000000", "VAR", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "269", "I32", "", "PF sum", "", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "271", "I32", "", "Phase sequence_1", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "273", "I32", "", "Hz", "", "1.000000", "0.000000", "Hz", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "275", "I32", "", "kWh(_) TOT", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "277", "I32", "", "kVarh(_) TOT", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "279", "I32", "", "kWh(-) TOT_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "287", "I32", "", "V L1-L2", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "289", "I32", "", "V L1-N", "", "0.100000", "0.000000", "V", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "291", "I32", "", "A L1", "", "0.001000", "0.000000", "A", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "293", "I32", "", "W L1", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "295", "I32", "", "VA L1", "", "0.100000", "0.000000", "VA", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "297", "I32", "", "VAR L1", "", "0.100000", "0.000000", "VAR", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "299", "I32", "", "PF L1", "", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "301", "I32", "", "V L2-L3", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "303", "I32", "", "V L2-N", "", "0.100000", "0.000000", "V", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "305", "I32", "", "A L2", "", "0.001000", "0.000000", "A", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "307", "I32", "", "W L2", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "309", "I32", "", "VA L2", "", "0.100000", "0.000000", "VA", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "311", "I32", "", "VAR L2", "", "0.100000", "0.000000", "VAR", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "313", "I32", "", "PF L2", "", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "315", "I32", "", "V L3-L1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "317", "I32", "", "V L3-N", "", "0.100000", "0.000000", "V", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "319", "I32", "", "A L3", "", "0.001000", "0.000000", "A", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "321", "I32", "", "W L3", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "323", "I32", "", "VA L3", "", "0.100000", "0.000000", "VA", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "325", "I32", "", "VAR L3", "", "0.100000", "0.000000", "VAR", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "327", "I32", "", "PF L3", "", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "385", "I32", "", "THD(15h) A L1_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "387", "I32", "", "THD(15h) V1-N_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "389", "I32", "", "THD(15h) V12_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "403", "I32", "", "THD(15h) A L2_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "405", "I32", "", "THD(15h) V2-N_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "407", "I32", "", "THD(15h) V23_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "421", "I32", "", "THD(15h) A L3_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "423", "I32", "", "THD(15h) V3-N_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "425", "I32", "", "THD(15h) V31_1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )

        # Section 2.5..2.8 Device Identification, Programming, and Diagnostics
        regs.append(
            ModbusRegisterDef(
                "3", "771", "U16", "", "Version code", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "772", "U16", "", "Revision code", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "773",
                "U16",
                "",
                "Programming lock (trimmer position in the rear of the display)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "12",
                "U16",
                "",
                "Carlo Gavazzi Controls identification code",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4099", "U16", "", "Measuring system", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4100", "U32", "", "Primary current", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "4102",
                "U32",
                "",
                "Voltage transformer ratio",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4104", "U16", "", "Sensor selection", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4105", "U16", "", "Range Rogowski sensor", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "4357",
                "U16",
                "",
                "Wrong connection (Installing help) management enabling",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "4358",
                "U16",
                "",
                "Wrong connection (installing help) status",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "4359",
                "U16",
                "",
                "THD calculation enabling",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4360", "I32", "", "I_counth", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "4115",
                "U16",
                "",
                "Pulse output duration (TON)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4129", "U16", "", "kWh per pulse", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4865", "U16", "", "Type of application", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "8193",
                "U16",
                "",
                "RS485 instrument address",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "8194", "U16", "", "RS485 baud rate", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "8195", "U16", "", "RS485 parity", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "8196", "U16", "", "RS485 Stop bit", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "16385",
                "U16",
                "",
                "Reset of all meters (kWhTOT_ kvarhTOT_ positive HoursCounter_ negative HoursCounter)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        # Serial Number (ASCII)
        regs.append(
            ModbusRegisterDef(
                "3",
                "20481",
                "U16",
                "",
                "SerialNumber_Letter 1 (from SX) Letter 2 (from SX)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "20482",
                "U16",
                "",
                "SerialNumber_Letter 3 (from SX) Letter 4 (from SX)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "20483",
                "U16",
                "",
                "SerialNumber_Letter 5 (from SX) Letter 6 (from SX)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "20484",
                "U16",
                "",
                "SerialNumber_Letter 7 (from SX) Letter 8 (from SX)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "20485",
                "U16",
                "",
                "SerialNumber_Letter 9 (from SX) Letter 10 (from SX)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "20486",
                "U16",
                "",
                "SerialNumber_Letter 11 (from SX) Letter 12 (from SX)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "20487",
                "U16",
                "",
                "SerialNumber_Letter 13 (from SX)",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "20488",
                "U16",
                "",
                "Instrument production year",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "20737",
                "U32",
                "",
                "Instrument secondary address",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )

        return regs

    # =========================================================================
    # 2. EM530 & EM540 (0-based physical addressing, exact 136 registers)
    # =========================================================================

    @classmethod
    def get_em500_registers(cls) -> list[ModbusRegisterDef]:
        """
        Build register definitions for Carlo Gavazzi EM530 and EM540.
        Source: EM530-EM540 COMMUNICATION PROTOCOL Rev. 1.3 & CarloGavazzi_EM530_test.csv.
        """
        regs: list[ModbusRegisterDef] = []

        # Instantaneous variables (Table 2.3)
        regs.append(
            ModbusRegisterDef(
                "3", "0", "I32", "", "V L1-N", "U_AC_1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "2", "I32", "", "V L2-N", "U_AC_2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4", "I32", "", "V L3-N", "U_AC_3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "6", "I32", "", "V L1-L2", "U_AC_L1L2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "8", "I32", "", "V L2-L3", "U_AC_L2L3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "10", "I32", "", "V L3-L1", "U_AC_L3L1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "12", "I32", "", "A L1", "I_AC_1", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "14", "I32", "", "A L2", "I_AC_2", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "16", "I32", "", "A L3", "I_AC_3", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "18", "I32", "", "W L1", "P_AC_1", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "20", "I32", "", "W L2", "P_AC_2", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "22", "I32", "", "W L3", "P_AC_3", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "24", "I32", "", "VA L1", "S_AC_1", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "26", "I32", "", "VA L2", "S_AC_2", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "28", "I32", "", "VA L3", "S_AC_3", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "30", "I32", "", "var L1", "Q_AC_1", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "32", "I32", "", "var L2", "Q_AC_2", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "34", "I32", "", "var L3", "Q_AC_3", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "36", "I32", "", "V L-N sys", "U_AC", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "38", "I32", "", "V L-L sys", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "40", "I32", "", "W sys", "P_AC", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "42", "I32", "", "VA sys", "S_AC", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "44", "I32", "", "var sys", "Q_AC", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "46", "I16", "", "PF L1", "PF_1", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "47", "I16", "", "PF L2", "PF_2", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "48", "I16", "", "PF L3", "PF_3", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "49", "I16", "", "PF sys", "PF", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "50", "I16", "", "Phase sequence", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "51", "I16", "", "Hz", "FREQ", "0.100000", "0.000000", "Hz", "4")
        )
        # Energies & Demands
        regs.append(
            ModbusRegisterDef(
                "3", "52", "I32", "", "kWh (+) TOT", "E_TOT", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "54", "I32", "", "kvarh (+) TOT", "Q_TOT", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "56", "I32", "", "kW dmd", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "58", "I32", "", "kW dmd peak", "", "0.100000", "0.000000", "kW", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "60", "I32", "", "kWh (+) PARTIAL", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "62", "I32", "", "kvarh (+) PARTIAL", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "64", "I32", "", "kWh (+) L1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "66", "I32", "", "kWh (+) L2", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "68", "I32", "", "kWh (+) L3", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "70", "I32", "", "kWh (+) t1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "72", "I32", "", "kWh (+) t2", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "74", "I32", "", "kWh (+) t3", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "76", "I32", "", "kWh (+) t4", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "78", "I32", "", "kWh (-) TOT", "E_EXP", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "80", "I32", "", "kvarh (-) TOT", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "82", "I32", "", "kWh (-) PARTIAL", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "84", "I32", "", "kvarh (-) PARTIAL", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "86", "I32", "", "kVAh TOT", "", "0.100000", "0.000000", "kVAh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "88", "I32", "", "kVAh PARTIAL", "", "0.100000", "0.000000", "kVAh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "90", "I32", "", "An", "I_AC_N", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "92", "I32", "", "Hour meter", "", "0.010000", "0.000000", "Hours", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "94", "I32", "", "Hour meter kWh (-)", "", "0.010000", "0.000000", "Hours", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "96", "I32", "", "kWh (-) L1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "98", "I32", "", "kWh (-) L2", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "100", "I32", "", "kWh (-) L3", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "102", "I32", "", "kWh (-) t1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "104", "I32", "", "kWh (-) t2", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "106", "I32", "", "kWh (-) t3", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "108", "I32", "", "kWh (-) t4", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "118", "I16", "", "PF sys_copy", "", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "120", "I16", "", "Hz_copy", "", "0.100000", "0.000000", "Hz", "4"
            )
        )
        # THD
        regs.append(
            ModbusRegisterDef(
                "3", "122", "I32", "", "THD A L1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "124", "I32", "", "THD A L2", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "126", "I32", "", "THD A L3", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "128", "I32", "", "THD V L1-N", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "130", "I32", "", "THD V L2-N", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "132", "I32", "", "THD V L3-N", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "134", "I32", "", "THD V L1-L2", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "136", "I32", "", "THD V L2-L3", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "138", "I32", "", "THD V L3-L1", "", "0.010000", "0.000000", "Percent", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "140", "I32", "", "A L1 dmd", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "142", "I32", "", "A L1 dmd peak", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "144", "I32", "", "A L2 dmd", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "146", "I32", "", "A L2 dmd peak", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "148", "I32", "", "A L3 dmd", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "150", "I32", "", "A L3 dmd peak", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "152", "I32", "", "An dmd", "", "0.001000", "0.000000", "A", "4")
        )

        # Grouped-by-phase duplicates (Section 2.4)
        regs.append(
            ModbusRegisterDef(
                "3",
                "248",
                "I32",
                "",
                "Hour meter kWh (-)_1",
                "",
                "0.010000",
                "0.000000",
                "Hours",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef("3", "254", "I32", "", "An_1", "", "0.001000", "0.000000", "A", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "258", "I32", "", "Hour meter_1", "", "0.010000", "0.000000", "Hours", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "260", "I32", "", "V L-N sys_1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "262", "I32", "", "V L-L sys_1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "264", "I32", "", "kW sys_1", "", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "266", "I32", "", "kVA sys_1", "", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "268", "I32", "", "kvar sys_1", "", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "270", "I32", "", "PF sys_1", "", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "272", "I32", "", "Phase sequence_1", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "274", "I32", "", "Hz_1", "", "0.100000", "0.000000", "Hz", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "276", "I32", "", "kWh (+) TOT_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "278", "I32", "", "Kvarh (+) TOT_1", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "280", "I32", "", "kWh (-) TOT_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "282", "I32", "", "kvarh (-) TOT_1", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "284", "I32", "", "W dmd_1", "", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "286", "I32", "", "W dmd peak_1", "", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "288", "I32", "", "V L1-L2_1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "290", "I32", "", "V L1-N_1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "292", "I32", "", "A L1_1", "", "0.001000", "0.000000", "A", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "294", "I32", "", "W L1_1", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "296", "I32", "", "VA L1_1", "", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "298", "I32", "", "var L1_1", "", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "300", "I32", "", "PF L1_1", "", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "302", "I32", "", "V L2-L3_1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "304", "I32", "", "V L2-N_1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "306", "I32", "", "A L2_1", "", "0.001000", "0.000000", "A", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "308", "I32", "", "W L2_1", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "310", "I32", "", "VA L2_1", "", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "312", "I32", "", "var L2_1", "", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "314", "I32", "", "PF L2_1", "", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "316", "I32", "", "V L3-L1_1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "318", "I32", "", "V L3-N_1", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "320", "I32", "", "A L3_1", "", "0.001000", "0.000000", "A", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "322", "I32", "", "W L3_1", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "324", "I32", "", "VA L3_1", "", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "326", "I32", "", "var L3_1", "", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "328", "I32", "", "PF L3_1", "", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "330", "I32", "", "kWh (+) PARTIAL_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "332",
                "I32",
                "",
                "Kvarh (+) PARTIAL_1",
                "",
                "0.100000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "334", "I32", "", "kWh (+) L1_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "336", "I32", "", "kWh (+) L2_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "338", "I32", "", "kWh (+) L3_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "340", "I32", "", "kWh (+) t1_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "342", "I32", "", "kWh (+) t2_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "344", "I32", "", "kWh (-) L1_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "346", "I32", "", "kWh (-) L2_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "348", "I32", "", "kWh (-) L3_1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )

        # High-resolution INT/DEC totalizers (Section 4.3)
        regs.append(
            ModbusRegisterDef(
                "3", "1024", "I32", "", "kWh (+) TOT - INT", "", "1.000000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1026", "I32", "", "kWh (+) TOT - DEC", "", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1028",
                "I32",
                "",
                "kvarh (+) TOT - INT",
                "",
                "1.000000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1030",
                "I32",
                "",
                "kvarh (+) TOT - DEC",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1032", "I32", "", "kWh (-) TOT - INT", "", "1.000000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1034", "I32", "", "kWh (-) TOT - DEC", "", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1036",
                "I32",
                "",
                "kvarh (-) TOT - INT",
                "",
                "1.000000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1038",
                "I32",
                "",
                "kvarh (-) TOT - DEC",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )

        # 64-bit integer totalizers (Section 4.4)
        regs.append(
            ModbusRegisterDef(
                "3",
                "1280",
                "I64",
                "",
                "kWh (+) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1284",
                "I64",
                "",
                "Kvarh (+) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1288",
                "I64",
                "",
                "kWh (+) Partial (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1292",
                "I64",
                "",
                "Kvarh (+) Partial (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1296",
                "I64",
                "",
                "kWh (+) L1 (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1300",
                "I64",
                "",
                "kWh (+) L2 (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1304",
                "I64",
                "",
                "kWh (+) L3 (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1308",
                "I64",
                "",
                "kWh (-) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1312",
                "I64",
                "",
                "kWh (-) Partial (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1316",
                "I64",
                "",
                "kvarh (-) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1320",
                "I64",
                "",
                "kvarh (-) Partial (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )

        return regs

    # =========================================================================
    # 3. EM100 Series: EM111, EM112, ET112 (1-Phase meters, 0-based addressing)
    # =========================================================================

    @classmethod
    def get_em100_registers(cls, model: str) -> list[ModbusRegisterDef]:
        """
        Build register definitions for Carlo Gavazzi EM111, EM112, and ET112.
        Source: EM100 Series Communication Protocol (em111_em112_et112_cp.pdf).
        """
        regs: list[ModbusRegisterDef] = []
        regs.append(
            ModbusRegisterDef(
                "3", "0", "I32", "", "V L-N", "U_AC_1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "2", "I32", "", "A", "I_AC_1", "0.001000", "0.000000", "A", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "4", "I32", "", "W", "P_AC", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "6", "I32", "", "VA", "S_AC", "0.100000", "0.000000", "VA", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "8", "I32", "", "var", "Q_AC", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "10", "I32", "", "W dmd", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "12", "I32", "", "W dmd peak", "", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "14", "I16", "", "PF", "PF", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef("3", "15", "I16", "", "Hz", "FREQ", "0.100000", "0.000000", "Hz", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "16", "I32", "", "kWh (+) TOT", "E_TOT", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "18", "I32", "", "Kvarh (+) TOT", "Q_TOT", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "20", "I32", "", "kWh (+) PARTIAL", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "22", "I32", "", "Kvarh (+) PARTIAL", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "24", "I32", "", "kWh (+) t1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "26", "I32", "", "kWh (+) t2", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "32", "I32", "", "kWh (-) TOT", "E_EXP", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "34", "I32", "", "kvarh (-) TOT", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "44", "I32", "", "Hour counter", "", "0.010000", "0.000000", "Hours", "4"
            )
        )

        # High-res and 64-bit energy counters
        regs.append(
            ModbusRegisterDef(
                "3", "1024", "I32", "", "kWh (+) TOT - INT", "", "1.000000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1026", "I32", "", "kWh (+) TOT - DEC", "", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1028",
                "I32",
                "",
                "kvarh (+) TOT - INT",
                "",
                "1.000000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1030",
                "I32",
                "",
                "kvarh (+) TOT - DEC",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1032", "I32", "", "kWh (-) TOT - INT", "", "1.000000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1034", "I32", "", "kWh (-) TOT - DEC", "", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1036",
                "I32",
                "",
                "kvarh (-) TOT - INT",
                "",
                "1.000000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1038",
                "I32",
                "",
                "kvarh (-) TOT - DEC",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )

        regs.append(
            ModbusRegisterDef(
                "3",
                "1280",
                "I64",
                "",
                "kWh (+) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1284",
                "I64",
                "",
                "Kvarh (+) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1308",
                "I64",
                "",
                "kWh (-) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1316",
                "I64",
                "",
                "kvarh (-) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )

        # Firmware & Controls ID
        regs.append(
            ModbusRegisterDef(
                "3", "771", "U16", "", "Version code", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "772", "U16", "", "Revision code", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "11",
                "U16",
                "",
                "Carlo Gavazzi Controls identification code",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        for s_idx in range(1, 8):
            regs.append(
                ModbusRegisterDef(
                    "3",
                    str(20480 + s_idx),
                    "U16",
                    "",
                    f"SerialNumber_Word_{s_idx}",
                    "",
                    "1.000000",
                    "0.000000",
                    "",
                    "4",
                )
            )
        regs.append(
            ModbusRegisterDef(
                "3", "20488", "U16", "", "Production year", "", "1.000000", "0.000000", "", "4"
            )
        )

        return regs

    # =========================================================================
    # 4. EM300 Series: EM330, EM340, ET330, ET340 (3-Phase meters, 0-based)
    # =========================================================================

    @classmethod
    def get_em300_registers(cls) -> list[ModbusRegisterDef]:
        """
        Build register definitions for Carlo Gavazzi EM330, EM340, ET330, and ET340.
        Source: EM300 Series Communication Protocol (em330_em340_et330_et340_cp.pdf).
        """
        regs: list[ModbusRegisterDef] = []
        regs.append(
            ModbusRegisterDef(
                "3", "0", "I32", "", "V L1-N", "U_AC_1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "2", "I32", "", "V L2-N", "U_AC_2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "4", "I32", "", "V L3-N", "U_AC_3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "6", "I32", "", "V L1-L2", "U_AC_L1L2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "8", "I32", "", "V L2-L3", "U_AC_L2L3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "10", "I32", "", "V L3-L1", "U_AC_L3L1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "12", "I32", "", "A L1", "I_AC_1", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "14", "I32", "", "A L2", "I_AC_2", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "16", "I32", "", "A L3", "I_AC_3", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "18", "I32", "", "W L1", "P_AC_1", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "20", "I32", "", "W L2", "P_AC_2", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "22", "I32", "", "W L3", "P_AC_3", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "24", "I32", "", "VA L1", "S_AC_1", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "26", "I32", "", "VA L2", "S_AC_2", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "28", "I32", "", "VA L3", "S_AC_3", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "30", "I32", "", "var L1", "Q_AC_1", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "32", "I32", "", "var L2", "Q_AC_2", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "34", "I32", "", "var L3", "Q_AC_3", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "36", "I32", "", "V L-N sys", "U_AC", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "38", "I32", "", "V L-L sys", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "40", "I32", "", "W sys", "P_AC", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "42", "I32", "", "VA sys", "S_AC", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "44", "I32", "", "var sys", "Q_AC", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "46", "I16", "", "PF L1", "PF_1", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "47", "I16", "", "PF L2", "PF_2", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "48", "I16", "", "PF L3", "PF_3", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "49", "I16", "", "PF sys", "PF", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "50", "I16", "", "Phase sequence", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "51", "I16", "", "Hz", "FREQ", "0.100000", "0.000000", "Hz", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "52", "I32", "", "kWh (+) TOT", "E_TOT", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "54", "I32", "", "Kvarh (+) TOT", "Q_TOT", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "56", "I32", "", "kW dmd", "", "0.100000", "0.000000", "W", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "58", "I32", "", "kW dmd peak", "", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "60", "I32", "", "kWh (+) PARTIAL", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "62", "I32", "", "Kvarh (+) PARTIAL", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "64", "I32", "", "kWh (+) L1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "66", "I32", "", "kWh (+) L2", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "68", "I32", "", "kWh (+) L3", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "70", "I32", "", "kWh (+) t1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "72", "I32", "", "kWh (+) t2", "", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "78", "I32", "", "kWh (-) TOT", "E_EXP", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "80", "I32", "", "kvarh (-) TOT", "", "0.100000", "0.000000", "kvarh", "4"
            )
        )

        # High resolution INT/DEC totalizers
        regs.append(
            ModbusRegisterDef(
                "3", "1024", "I32", "", "kWh (+) TOT - INT", "", "1.000000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1026", "I32", "", "kWh (+) TOT - DEC", "", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1028",
                "I32",
                "",
                "kvarh (+) TOT - INT",
                "",
                "1.000000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1030",
                "I32",
                "",
                "kvarh (+) TOT - DEC",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1032", "I32", "", "kWh (-) TOT - INT", "", "1.000000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1034", "I32", "", "kWh (-) TOT - DEC", "", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1036",
                "I32",
                "",
                "kvarh (-) TOT - INT",
                "",
                "1.000000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1038",
                "I32",
                "",
                "kvarh (-) TOT - DEC",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )

        # 64-bit totalizers
        regs.append(
            ModbusRegisterDef(
                "3",
                "1280",
                "I64",
                "",
                "kWh (+) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1284",
                "I64",
                "",
                "Kvarh (+) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1308",
                "I64",
                "",
                "kWh (-) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kWh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1316",
                "I64",
                "",
                "kvarh (-) TOT (64-bit)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )

        # Device ID
        regs.append(
            ModbusRegisterDef(
                "3", "771", "U16", "", "Version code", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "772", "U16", "", "Revision code", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "11",
                "U16",
                "",
                "Carlo Gavazzi Controls identification code",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )
        for s_idx in range(1, 8):
            regs.append(
                ModbusRegisterDef(
                    "3",
                    str(20480 + s_idx),
                    "U16",
                    "",
                    f"SerialNumber_Word_{s_idx}",
                    "",
                    "1.000000",
                    "0.000000",
                    "",
                    "4",
                )
            )
        regs.append(
            ModbusRegisterDef(
                "3", "20488", "U16", "", "Production year", "", "1.000000", "0.000000", "", "4"
            )
        )

        return regs

    # =========================================================================
    # 5. EM21 (3-Phase meter, 1-based addressing)
    # =========================================================================

    @classmethod
    def get_em21_registers(cls) -> list[ModbusRegisterDef]:
        """
        Build register definitions for Carlo Gavazzi EM21.
        Source: EM21 COMMUNICATION PROTOCOL V1r2 2010.pdf.
        """
        regs: list[ModbusRegisterDef] = []
        regs.append(
            ModbusRegisterDef(
                "3", "1", "I32", "", "V L1-N", "U_AC_1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "3", "I32", "", "V L2-N", "U_AC_2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "5", "I32", "", "V L3-N", "U_AC_3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "7", "I32", "", "V L1-L2", "U_AC_L1L2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "9", "I32", "", "V L2-L3", "U_AC_L2L3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "11", "I32", "", "V L3-L1", "U_AC_L3L1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "13", "I32", "", "A L1", "I_AC_1", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "15", "I32", "", "A L2", "I_AC_2", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "17", "I32", "", "A L3", "I_AC_3", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "19", "I32", "", "W L1", "P_AC_1", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "21", "I32", "", "W L2", "P_AC_2", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "23", "I32", "", "W L3", "P_AC_3", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "25", "I32", "", "VA L1", "S_AC_1", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "27", "I32", "", "VA L2", "S_AC_2", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "29", "I32", "", "VA L3", "S_AC_3", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "31", "I32", "", "VAR L1", "Q_AC_1", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "33", "I32", "", "VAR L2", "Q_AC_2", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "35", "I32", "", "VAR L3", "Q_AC_3", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "37", "I32", "", "V L-N sys", "U_AC", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "39", "I32", "", "V L-L sys", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "41", "I32", "", "W sys", "P_AC", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "43", "I32", "", "VA sys", "S_AC", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "45", "I32", "", "VAR sys", "Q_AC", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "47", "I16", "", "PF L1", "PF_1", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "48", "I16", "", "PF L2", "PF_2", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "49", "I16", "", "PF L3", "PF_3", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "50", "I16", "", "PF sys", "PF", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "51", "I16", "", "Phase sequence", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "52", "I16", "", "Hz", "FREQ", "1.000000", "0.000000", "Hz", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "53", "I32", "", "kWh(+) TOT", "E_TOT", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "55", "I32", "", "kvarh(+) TOT", "Q_TOT", "0.100000", "0.000000", "kvarh", "4"
            )
        )

        regs.append(
            ModbusRegisterDef(
                "3", "771", "U16", "", "Version code", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "772", "U16", "", "Revision code", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "773", "U16", "", "Programming lock", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "12",
                "U16",
                "",
                "Carlo Gavazzi Controls identification code",
                "",
                "1.000000",
                "0.000000",
                "",
                "4",
            )
        )

        return regs

    # =========================================================================
    # 6. EM271 (Dual 3-Phase Multi-Circuit Energy Meter)
    # =========================================================================

    @classmethod
    def get_em271_registers(cls) -> list[ModbusRegisterDef]:
        """
        Build register definitions for Carlo Gavazzi EM271.
        Source: EM271 CP V0R0 ENG.PDF.
        """
        regs: list[ModbusRegisterDef] = []
        regs.append(
            ModbusRegisterDef(
                "3", "1", "I32", "", "V L1-N", "U_AC_1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "3", "I32", "", "V L2-N", "U_AC_2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "5", "I32", "", "V L3-N", "U_AC_3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "7", "I32", "", "V L1-L2", "U_AC_L1L2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "9", "I32", "", "V L2-L3", "U_AC_L2L3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "11", "I32", "", "V L3-L1", "U_AC_L3L1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "13", "I32", "", "A L1 sum", "I_AC_1", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "15", "I32", "", "A L2 sum", "I_AC_2", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "17", "I32", "", "A L3 sum", "I_AC_3", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "19", "I32", "", "W sum", "P_AC", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "21", "I32", "", "VA sum", "S_AC", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "23", "I32", "", "VAR sum", "Q_AC", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "25", "I32", "", "kWh(+) TOT sum", "E_TOT", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "27",
                "I32",
                "",
                "kvarh(+) TOT sum",
                "Q_TOT",
                "0.100000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef("3", "29", "I16", "", "Hz", "FREQ", "0.100000", "0.000000", "Hz", "4")
        )

        # Channel 1 and Channel 2 distinct loads
        regs.append(
            ModbusRegisterDef(
                "3", "41", "I32", "", "A L1 Ch1", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "43", "I32", "", "A L2 Ch1", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "45", "I32", "", "A L3 Ch1", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "47", "I32", "", "W Ch1", "P_AC_1", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "55", "I32", "", "kWh(+) TOT Ch1", "", "0.100000", "0.000000", "kWh", "4"
            )
        )

        regs.append(
            ModbusRegisterDef(
                "3", "61", "I32", "", "A L1 Ch2", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "63", "I32", "", "A L2 Ch2", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "65", "I32", "", "A L3 Ch2", "", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "67", "I32", "", "W Ch2", "P_AC_2", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "75", "I32", "", "kWh(+) TOT Ch2", "", "0.100000", "0.000000", "kWh", "4"
            )
        )

        return regs

    # =========================================================================
    # 7. WM15 (3-Phase Power Analyzer)
    # =========================================================================

    @classmethod
    def get_wm15_registers(cls) -> list[ModbusRegisterDef]:
        """
        Build register definitions for Carlo Gavazzi WM15 power analyzer.
        Source: WM15-modbus_registers-V1-1-20192510.pdf.
        """
        regs: list[ModbusRegisterDef] = []
        regs.append(
            ModbusRegisterDef(
                "3", "1", "I32", "", "V L1-N", "U_AC_1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "3", "I32", "", "V L2-N", "U_AC_2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "5", "I32", "", "V L3-N", "U_AC_3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "7", "I32", "", "V L1-L2", "U_AC_L1L2", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "9", "I32", "", "V L2-L3", "U_AC_L2L3", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "11", "I32", "", "V L3-L1", "U_AC_L3L1", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "13", "I32", "", "A L1", "I_AC_1", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "15", "I32", "", "A L2", "I_AC_2", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "17", "I32", "", "A L3", "I_AC_3", "0.001000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "19", "I32", "", "W L1", "P_AC_1", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "21", "I32", "", "W L2", "P_AC_2", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "23", "I32", "", "W L3", "P_AC_3", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "25", "I32", "", "VA L1", "S_AC_1", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "27", "I32", "", "VA L2", "S_AC_2", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "29", "I32", "", "VA L3", "S_AC_3", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "31", "I32", "", "VAR L1", "Q_AC_1", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "33", "I32", "", "VAR L2", "Q_AC_2", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "35", "I32", "", "VAR L3", "Q_AC_3", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "37", "I32", "", "V L-N sys", "U_AC", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "39", "I32", "", "V L-L sys", "", "0.100000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "41", "I32", "", "W sys", "P_AC", "0.100000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "43", "I32", "", "VA sys", "S_AC", "0.100000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "45", "I32", "", "VAR sys", "Q_AC", "0.100000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "47", "I16", "", "PF L1", "PF_1", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "48", "I16", "", "PF L2", "PF_2", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "49", "I16", "", "PF L3", "PF_3", "0.001000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "50", "I16", "", "PF sys", "PF", "0.001000", "0.000000", "", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "51", "I16", "", "Phase sequence", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef("3", "52", "I16", "", "Hz", "FREQ", "0.100000", "0.000000", "Hz", "4")
        )
        regs.append(
            ModbusRegisterDef(
                "3", "53", "I32", "", "kWh(+) TOT", "E_TOT", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "55", "I32", "", "kvarh(+) TOT", "Q_TOT", "0.100000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "79", "I32", "", "kWh(-) TOT", "E_EXP", "0.100000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "153",
                "I32",
                "",
                "Neutral current (In)",
                "I_AC_N",
                "0.001000",
                "0.000000",
                "A",
                "4",
            )
        )

        return regs

    # =========================================================================
    # 8. WM20 & WM50 (Modular Power Analyzers, IEEE 754 Float format)
    # =========================================================================

    @classmethod
    def get_wm20_wm50_registers(cls, model: str) -> list[ModbusRegisterDef]:
        """
        Build register definitions for Carlo Gavazzi WM20 and WM50.
        Sources: WM20-modbus_registers.pdf & wm50_modbus_registers.pdf.
        """
        regs: list[ModbusRegisterDef] = []
        # IEEE 754 Real-Time Instantaneous Variables (FLOAT / F32)
        regs.append(
            ModbusRegisterDef(
                "3", "81", "FLOAT", "", "V L1-N", "U_AC_1", "1.000000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "83", "FLOAT", "", "V L2-N", "U_AC_2", "1.000000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "85", "FLOAT", "", "V L3-N", "U_AC_3", "1.000000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "87", "FLOAT", "", "V L-N sys", "U_AC", "1.000000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "89", "FLOAT", "", "V L1-L2", "U_AC_L1L2", "1.000000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "91", "FLOAT", "", "V L2-L3", "U_AC_L2L3", "1.000000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "93", "FLOAT", "", "V L3-L1", "U_AC_L3L1", "1.000000", "0.000000", "V", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "95", "FLOAT", "", "V L-L sys", "", "1.000000", "0.000000", "V", "4"
            )
        )

        regs.append(
            ModbusRegisterDef(
                "3", "97", "FLOAT", "", "A L1", "I_AC_1", "1.000000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "99", "FLOAT", "", "A L2", "I_AC_2", "1.000000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "101", "FLOAT", "", "A L3", "I_AC_3", "1.000000", "0.000000", "A", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "103", "FLOAT", "", "A N", "I_AC_N", "1.000000", "0.000000", "A", "4"
            )
        )

        regs.append(
            ModbusRegisterDef(
                "3", "105", "FLOAT", "", "W L1", "P_AC_1", "1.000000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "107", "FLOAT", "", "W L2", "P_AC_2", "1.000000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "109", "FLOAT", "", "W L3", "P_AC_3", "1.000000", "0.000000", "W", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "111", "FLOAT", "", "W sys", "P_AC", "1.000000", "0.000000", "W", "4"
            )
        )

        regs.append(
            ModbusRegisterDef(
                "3", "113", "FLOAT", "", "VA L1", "S_AC_1", "1.000000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "115", "FLOAT", "", "VA L2", "S_AC_2", "1.000000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "117", "FLOAT", "", "VA L3", "S_AC_3", "1.000000", "0.000000", "VA", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "119", "FLOAT", "", "VA sys", "S_AC", "1.000000", "0.000000", "VA", "4"
            )
        )

        regs.append(
            ModbusRegisterDef(
                "3", "121", "FLOAT", "", "VAR L1", "Q_AC_1", "1.000000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "123", "FLOAT", "", "VAR L2", "Q_AC_2", "1.000000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "125", "FLOAT", "", "VAR L3", "Q_AC_3", "1.000000", "0.000000", "var", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "127", "FLOAT", "", "VAR sys", "Q_AC", "1.000000", "0.000000", "var", "4"
            )
        )

        regs.append(
            ModbusRegisterDef(
                "3", "129", "FLOAT", "", "PF L1", "PF_1", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "131", "FLOAT", "", "PF L2", "PF_2", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "133", "FLOAT", "", "PF L3", "PF_3", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "135", "FLOAT", "", "PF sys", "PF", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "137", "FLOAT", "", "Hz", "FREQ", "1.000000", "0.000000", "Hz", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "143", "FLOAT", "", "Phase sequence", "", "1.000000", "0.000000", "", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "145", "FLOAT", "", "A sys", "", "1.000000", "0.000000", "A", "4"
            )
        )

        # Energy Totalizers (UINT64, in Wh or varh, converted with 0.001 to kWh / kvarh)
        regs.append(
            ModbusRegisterDef(
                "3", "1281", "U64", "", "Total kWh (+)", "E_TOT", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1285",
                "U64",
                "",
                "Total kvarh (+)",
                "Q_TOT",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1289", "U64", "", "Total kWh (-)", "E_EXP", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1293", "U64", "", "Total kvarh (-)", "", "0.001000", "0.000000", "kvarh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1297", "U64", "", "Partial kWh (+)", "", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1301",
                "U64",
                "",
                "Partial kvarh (+)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1305", "U64", "", "Partial kWh (-)", "", "0.001000", "0.000000", "kWh", "4"
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3",
                "1309",
                "U64",
                "",
                "Partial kvarh (-)",
                "",
                "0.001000",
                "0.000000",
                "kvarh",
                "4",
            )
        )
        regs.append(
            ModbusRegisterDef(
                "3", "1313", "U64", "", "Hours counter", "", "0.010000", "0.000000", "Hours", "4"
            )
        )

        return regs

    # =========================================================================
    # Master Generation Dispatcher
    # =========================================================================

    @classmethod
    def get_model_definitions(cls) -> dict[str, list[ModbusRegisterDef]]:
        """Return all certified Carlo Gavazzi models mapped to their register definitions."""
        return {
            "EM210_MV": cls.get_em210_registers(is_mv=True),
            "EM210": cls.get_em210_registers(is_mv=False),
            "EM530": cls.get_em500_registers(),
            "EM540": cls.get_em500_registers(),
            "EM111": cls.get_em100_registers("EM111"),
            "EM112": cls.get_em100_registers("EM112"),
            "ET112": cls.get_em100_registers("ET112"),
            "EM330": cls.get_em300_registers(),
            "EM340": cls.get_em300_registers(),
            "ET330": cls.get_em300_registers(),
            "ET340": cls.get_em300_registers(),
            "EM21": cls.get_em21_registers(),
            "EM271": cls.get_em271_registers(),
            "WM15": cls.get_wm15_registers(),
            "WM20": cls.get_wm20_wm50_registers("WM20"),
            "WM50": cls.get_wm20_wm50_registers("WM50"),
        }

    @classmethod
    def generate_all(cls, repo_root: Path) -> dict[str, Path]:
        """
        Generate and synchronize all Carlo Gavazzi definitions across the repository.
        Returns a mapping of model name to generated production file path.
        """
        models = cls.get_model_definitions()
        output_paths: dict[str, Path] = {}

        prod_dir = repo_root / "production"
        def_dir = repo_root / "Definitions"
        cg_def_dir = repo_root / "Definitions" / "CarloGavazzi"
        gen_dir = repo_root / "generated_definitions"
        tag_dir = repo_root / "tagged_definitions"

        target_dirs = [prod_dir, def_dir, cg_def_dir, gen_dir, tag_dir]
        for d in target_dirs:
            d.mkdir(parents=True, exist_ok=True)

        for model, regs in models.items():
            filename = f"MET_CarloGavazzi_{model}.csv"
            prod_file = prod_dir / filename
            cls.write_definition_file(prod_file, model, regs)
            output_paths[model] = prod_file

            # Mirror to target synchronization directories
            for target_dir in [cg_def_dir, def_dir, gen_dir, tag_dir]:
                mirror_file = target_dir / filename
                mirror_file.write_bytes(prod_file.read_bytes())

        return output_paths
