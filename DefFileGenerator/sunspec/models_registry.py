"""
Official SunSpec Alliance Information Models Registry.

Provides typed definitions of official SunSpec models, register offsets,
primitive types, register counts, units, scale factor associations,
and canonical SunPM telemetry tags.

References:
- SunSpec Inverter Models: 1, 101, 102, 103, 120, 121, 122, 123, 124, 126, 160
- SunSpec Meter Models: 201, 202, 203, 204
- SunSpec DER 700 Series: 701, 702, 704, 714
- SunSpec Storage Models: 802
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class SunSpecPoint:
    """Represents a single data point in a SunSpec model."""
    offset: int                    # Offset in 16-bit registers from the start of model payload (after ID and L)
    symbol: str                    # SunSpec symbol (e.g. "A", "W", "A_SF", "Hz")
    data_type: str                 # SunSpec type (uint16, int16, uint32, int32, uint64, float32, string, sunssf, bitfield16, bitfield32, enum16, pad)
    register_count: int            # Word count (1 for 16-bit, 2 for 32-bit, 4 for 64-bit, 8/16 for strings)
    scale_factor_point: str | None = None  # Symbol of referenced sunssf point (e.g. "A_SF")
    unit: str = ""                 # Engineering unit (A, V, W, Hz, degC, Wh, %, etc.)
    access: str = "R"              # "R" (read-only) or "RW" (read-write)
    action_code: str = "4"         # 4: Measurement, 3: Param, 1: Alarm, 2: Event
    sunpm_tag: str | None = None   # Canonical Webdyn SunPM tag
    description: str = ""          # Human-readable variable label

    @property
    def is_pad(self) -> bool:
        return self.data_type.lower() == "pad" or self.symbol.lower().endswith("pad")

    @property
    def is_scale_factor(self) -> bool:
        return self.data_type.lower() == "sunssf" or self.symbol.endswith("_SF")


@dataclass
class SunSpecModelDef:
    """Represents an official SunSpec model definition."""
    model_id: int
    name: str
    description: str
    standard_length: int
    points: list[SunSpecPoint]
    is_variable_length: bool = False
    repeating_block_length: int = 0
    repeating_points_template: list[SunSpecPoint] = field(default_factory=list)

    def get_points(self, num_repeats: int = 0, l_common: int = 65) -> list[SunSpecPoint]:
        """
        Returns the resolved list of points for this model.
        For variable models (e.g. 160 or 714), expands repeating blocks.
        For Common Model 1, adjusts points based on actual length L (65 vs 66).
        """
        result = list(self.points)

        # Handle Common Model 1 trailing pad register when L=66
        if self.model_id == 1 and l_common == 66:
            # Add pad register at offset 65 if not already present
            if not any(p.symbol == "Pad" for p in result):
                result.append(
                    SunSpecPoint(
                        offset=65,
                        symbol="Pad",
                        data_type="pad",
                        register_count=1,
                        description="SunSpec Common Model 64-bit Alignment Pad",
                    )
                )

        if self.is_variable_length and num_repeats > 0 and self.repeating_points_template:
            base_offset = self.standard_length
            for r in range(1, num_repeats + 1):
                block_offset = base_offset + (r - 1) * self.repeating_block_length
                for pt in self.repeating_points_template:
                    tag = pt.sunpm_tag
                    if tag and "{i}" in tag:
                        tag = tag.format(i=r)
                    elif tag and not tag.endswith(f"_{r}") and tag not in ("MANUFACTURER", "MODEL", "SERIAL"):
                        tag = f"{tag}_{r}"

                    desc = pt.description
                    if "{i}" in desc:
                        desc = desc.format(i=r)
                    else:
                        desc = f"{desc} #{r}"

                    expanded_pt = SunSpecPoint(
                        offset=block_offset + pt.offset,
                        symbol=f"{pt.symbol}_{r}",
                        data_type=pt.data_type,
                        register_count=pt.register_count,
                        scale_factor_point=pt.scale_factor_point,
                        unit=pt.unit,
                        access=pt.access,
                        action_code=pt.action_code,
                        sunpm_tag=tag,
                        description=desc,
                    )
                    result.append(expanded_pt)

        return result


# ==============================================================================
# MODEL 1 : Common Model (Device Information)
# ==============================================================================
MODEL_1_POINTS = [
    SunSpecPoint(0, "Mn", "string16", 16, description="SunSpec Manufacturer Name", sunpm_tag="MANUFACTURER"),
    SunSpecPoint(16, "Md", "string16", 16, description="SunSpec Device Model", sunpm_tag="MODEL"),
    SunSpecPoint(32, "Opt", "string8", 8, description="SunSpec Device Options"),
    SunSpecPoint(40, "Vr", "string8", 8, description="SunSpec Firmware Version", sunpm_tag="FIRMWARE"),
    SunSpecPoint(48, "SN", "string16", 16, description="SunSpec Serial Number", sunpm_tag="SERIAL"),
    SunSpecPoint(64, "DA", "uint16", 1, access="RW", action_code="8", description="SunSpec Modbus Device Address"),
]

# ==============================================================================
# MODEL 101 : Single Phase Inverter (AC Measurements + DC + Status)
# ==============================================================================
MODEL_101_POINTS = [
    SunSpecPoint(0, "A", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC", description="AC Current Total"),
    SunSpecPoint(1, "AphA", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_1", description="AC Current Phase A"),
    SunSpecPoint(2, "AphB", "uint16", 1, scale_factor_point="A_SF", unit="A", description="AC Current Phase B (Unused 1-Ph)"),
    SunSpecPoint(3, "AphC", "uint16", 1, scale_factor_point="A_SF", unit="A", description="AC Current Phase C (Unused 1-Ph)"),
    SunSpecPoint(4, "A_SF", "sunssf", 1, description="Current Scale Factor"),
    SunSpecPoint(5, "PPVphAB", "uint16", 1, scale_factor_point="V_SF", unit="V", description="Phase-to-Phase Voltage AB"),
    SunSpecPoint(6, "PPVphBC", "uint16", 1, scale_factor_point="V_SF", unit="V", description="Phase-to-Phase Voltage BC"),
    SunSpecPoint(7, "PPVphCA", "uint16", 1, scale_factor_point="V_SF", unit="V", description="Phase-to-Phase Voltage CA"),
    SunSpecPoint(8, "PhVphA", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_1", description="Phase-to-Neutral Voltage A"),
    SunSpecPoint(9, "PhVphB", "uint16", 1, scale_factor_point="V_SF", unit="V", description="Phase-to-Neutral Voltage B"),
    SunSpecPoint(10, "PhVphC", "uint16", 1, scale_factor_point="V_SF", unit="V", description="Phase-to-Neutral Voltage C"),
    SunSpecPoint(11, "V_SF", "sunssf", 1, description="Voltage Scale Factor"),
    SunSpecPoint(12, "W", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC", description="AC Active Power"),
    SunSpecPoint(13, "W_SF", "sunssf", 1, description="Power Scale Factor"),
    SunSpecPoint(14, "Hz", "uint16", 1, scale_factor_point="Hz_SF", unit="Hz", sunpm_tag="FREQ", description="Line Frequency"),
    SunSpecPoint(15, "Hz_SF", "sunssf", 1, description="Frequency Scale Factor"),
    SunSpecPoint(16, "VA", "int16", 1, scale_factor_point="VA_SF", unit="VA", sunpm_tag="S_AC", description="Apparent Power"),
    SunSpecPoint(17, "VA_SF", "sunssf", 1, description="Apparent Power Scale Factor"),
    SunSpecPoint(18, "VAr", "int16", 1, scale_factor_point="VAr_SF", unit="var", sunpm_tag="Q_AC", description="Reactive Power"),
    SunSpecPoint(19, "VAr_SF", "sunssf", 1, description="Reactive Power Scale Factor"),
    SunSpecPoint(20, "PF", "int16", 1, scale_factor_point="PF_SF", unit="pct", sunpm_tag="PF", description="Power Factor"),
    SunSpecPoint(21, "PF_SF", "sunssf", 1, description="Power Factor Scale Factor"),
    SunSpecPoint(22, "WH", "uint32", 2, scale_factor_point="WH_SF", unit="Wh", sunpm_tag="E_TOT", description="Lifetime Energy Production"),
    SunSpecPoint(24, "WH_SF", "sunssf", 1, description="Energy Scale Factor"),
    SunSpecPoint(25, "DCA", "uint16", 1, scale_factor_point="DCA_SF", unit="A", sunpm_tag="I_DC", description="DC Current"),
    SunSpecPoint(26, "DCA_SF", "sunssf", 1, description="DC Current Scale Factor"),
    SunSpecPoint(27, "DCV", "uint16", 1, scale_factor_point="DCV_SF", unit="V", sunpm_tag="U_DC", description="DC Voltage"),
    SunSpecPoint(28, "DCV_SF", "sunssf", 1, description="DC Voltage Scale Factor"),
    SunSpecPoint(29, "DCW", "int16", 1, scale_factor_point="DCW_SF", unit="W", sunpm_tag="P_DC", description="DC Power"),
    SunSpecPoint(30, "DCW_SF", "sunssf", 1, description="DC Power Scale Factor"),
    SunSpecPoint(31, "TmpCab", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", sunpm_tag="TEMP_INT", description="Cabinet Temperature"),
    SunSpecPoint(32, "TmpSnk", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", sunpm_tag="TEMP_EXT", description="Heat Sink Temperature"),
    SunSpecPoint(33, "TmpTrns", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", description="Transformer Temperature"),
    SunSpecPoint(34, "TmpOt", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", description="Other Temperature"),
    SunSpecPoint(35, "Tmp_SF", "sunssf", 1, description="Temperature Scale Factor"),
    SunSpecPoint(36, "St", "enum16", 1, sunpm_tag="STATE", description="Inverter Operating State"),
    SunSpecPoint(37, "StVnd", "enum16", 1, description="Vendor Specific Operating State"),
    SunSpecPoint(38, "Evt1", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="Inverter Event 1 / Alarms"),
    SunSpecPoint(40, "Evt2", "bitfield32", 2, action_code="1", description="Inverter Event 2 / Warnings"),
    SunSpecPoint(42, "EvtVnd1", "bitfield32", 2, action_code="1", description="Vendor Event 1"),
    SunSpecPoint(44, "EvtVnd2", "bitfield32", 2, action_code="1", description="Vendor Event 2"),
    SunSpecPoint(46, "EvtVnd3", "bitfield32", 2, action_code="1", description="Vendor Event 3"),
    SunSpecPoint(48, "EvtVnd4", "bitfield32", 2, action_code="1", description="Vendor Event 4"),
]

# ==============================================================================
# MODEL 102 : Split Phase Inverter (AC Measurements + DC + Status)
# ==============================================================================
MODEL_102_POINTS = [
    SunSpecPoint(0, "A", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC", description="AC Current Total"),
    SunSpecPoint(1, "AphA", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_1", description="AC Current Phase A"),
    SunSpecPoint(2, "AphB", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_2", description="AC Current Phase B"),
    SunSpecPoint(3, "AphC", "uint16", 1, scale_factor_point="A_SF", unit="A", description="AC Current Phase C (Unused Split-Phase)"),
    SunSpecPoint(4, "A_SF", "sunssf", 1, description="Current Scale Factor"),
    SunSpecPoint(5, "PPVphAB", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L1L2", description="Phase-to-Phase Voltage AB"),
    SunSpecPoint(6, "PPVphBC", "uint16", 1, scale_factor_point="V_SF", unit="V", description="Phase-to-Phase Voltage BC (Unused)"),
    SunSpecPoint(7, "PPVphCA", "uint16", 1, scale_factor_point="V_SF", unit="V", description="Phase-to-Phase Voltage CA (Unused)"),
    SunSpecPoint(8, "PhVphA", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_1", description="Phase-to-Neutral Voltage A"),
    SunSpecPoint(9, "PhVphB", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_2", description="Phase-to-Neutral Voltage B"),
    SunSpecPoint(10, "PhVphC", "uint16", 1, scale_factor_point="V_SF", unit="V", description="Phase-to-Neutral Voltage C (Unused)"),
    SunSpecPoint(11, "V_SF", "sunssf", 1, description="Voltage Scale Factor"),
    SunSpecPoint(12, "W", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC", description="AC Active Power"),
    SunSpecPoint(13, "W_SF", "sunssf", 1, description="Power Scale Factor"),
    SunSpecPoint(14, "Hz", "uint16", 1, scale_factor_point="Hz_SF", unit="Hz", sunpm_tag="FREQ", description="Line Frequency"),
    SunSpecPoint(15, "Hz_SF", "sunssf", 1, description="Frequency Scale Factor"),
    SunSpecPoint(16, "VA", "int16", 1, scale_factor_point="VA_SF", unit="VA", sunpm_tag="S_AC", description="Apparent Power"),
    SunSpecPoint(17, "VA_SF", "sunssf", 1, description="Apparent Power Scale Factor"),
    SunSpecPoint(18, "VAr", "int16", 1, scale_factor_point="VAr_SF", unit="var", sunpm_tag="Q_AC", description="Reactive Power"),
    SunSpecPoint(19, "VAr_SF", "sunssf", 1, description="Reactive Power Scale Factor"),
    SunSpecPoint(20, "PF", "int16", 1, scale_factor_point="PF_SF", unit="pct", sunpm_tag="PF", description="Power Factor"),
    SunSpecPoint(21, "PF_SF", "sunssf", 1, description="Power Factor Scale Factor"),
    SunSpecPoint(22, "WH", "uint32", 2, scale_factor_point="WH_SF", unit="Wh", sunpm_tag="E_TOT", description="Lifetime Energy Production"),
    SunSpecPoint(24, "WH_SF", "sunssf", 1, description="Energy Scale Factor"),
    SunSpecPoint(25, "DCA", "uint16", 1, scale_factor_point="DCA_SF", unit="A", sunpm_tag="I_DC", description="DC Current Total"),
    SunSpecPoint(26, "DCA_SF", "sunssf", 1, description="DC Current Scale Factor"),
    SunSpecPoint(27, "DCV", "uint16", 1, scale_factor_point="DCV_SF", unit="V", sunpm_tag="U_DC", description="DC Voltage Total"),
    SunSpecPoint(28, "DCV_SF", "sunssf", 1, description="DC Voltage Scale Factor"),
    SunSpecPoint(29, "DCW", "int16", 1, scale_factor_point="DCW_SF", unit="W", sunpm_tag="P_DC", description="DC Power Total"),
    SunSpecPoint(30, "DCW_SF", "sunssf", 1, description="DC Power Scale Factor"),
    SunSpecPoint(31, "TmpCab", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", sunpm_tag="TEMP_INT", description="Cabinet Temperature"),
    SunSpecPoint(32, "TmpSnk", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", sunpm_tag="TEMP_EXT", description="Heat Sink Temperature"),
    SunSpecPoint(33, "TmpTrns", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", description="Transformer Temperature"),
    SunSpecPoint(34, "TmpOt", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", description="Other Temperature"),
    SunSpecPoint(35, "Tmp_SF", "sunssf", 1, description="Temperature Scale Factor"),
    SunSpecPoint(36, "St", "enum16", 1, sunpm_tag="STATE", description="Inverter Operating State"),
    SunSpecPoint(37, "StVnd", "enum16", 1, description="Vendor Specific Operating State"),
    SunSpecPoint(38, "Evt1", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="Inverter Event 1 / Alarms"),
    SunSpecPoint(40, "Evt2", "bitfield32", 2, action_code="1", description="Inverter Event 2 / Warnings"),
    SunSpecPoint(42, "EvtVnd1", "bitfield32", 2, action_code="1", description="Vendor Event 1"),
    SunSpecPoint(44, "EvtVnd2", "bitfield32", 2, action_code="1", description="Vendor Event 2"),
    SunSpecPoint(46, "EvtVnd3", "bitfield32", 2, action_code="1", description="Vendor Event 3"),
    SunSpecPoint(48, "EvtVnd4", "bitfield32", 2, action_code="1", description="Vendor Event 4"),
]

# ==============================================================================
# MODEL 103 : Three Phase Inverter (AC Measurements + DC + Status)
# ==============================================================================
MODEL_103_POINTS = [
    SunSpecPoint(0, "A", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC", description="AC Current Total"),
    SunSpecPoint(1, "AphA", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_1", description="AC Current Phase A"),
    SunSpecPoint(2, "AphB", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_2", description="AC Current Phase B"),
    SunSpecPoint(3, "AphC", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_3", description="AC Current Phase C"),
    SunSpecPoint(4, "A_SF", "sunssf", 1, description="Current Scale Factor"),
    SunSpecPoint(5, "PPVphAB", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L1L2", description="Phase Voltage AB"),
    SunSpecPoint(6, "PPVphBC", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L2L3", description="Phase Voltage BC"),
    SunSpecPoint(7, "PPVphCA", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L3L1", description="Phase Voltage CA"),
    SunSpecPoint(8, "PhVphA", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_1", description="Phase Voltage A-N"),
    SunSpecPoint(9, "PhVphB", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_2", description="Phase Voltage B-N"),
    SunSpecPoint(10, "PhVphC", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_3", description="Phase Voltage C-N"),
    SunSpecPoint(11, "V_SF", "sunssf", 1, description="Voltage Scale Factor"),
    SunSpecPoint(12, "W", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC", description="AC Active Power"),
    SunSpecPoint(13, "W_SF", "sunssf", 1, description="Power Scale Factor"),
    SunSpecPoint(14, "Hz", "uint16", 1, scale_factor_point="Hz_SF", unit="Hz", sunpm_tag="FREQ", description="Line Frequency"),
    SunSpecPoint(15, "Hz_SF", "sunssf", 1, description="Frequency Scale Factor"),
    SunSpecPoint(16, "VA", "int16", 1, scale_factor_point="VA_SF", unit="VA", sunpm_tag="S_AC", description="Apparent Power"),
    SunSpecPoint(17, "VA_SF", "sunssf", 1, description="Apparent Power Scale Factor"),
    SunSpecPoint(18, "VAr", "int16", 1, scale_factor_point="VAr_SF", unit="var", sunpm_tag="Q_AC", description="Reactive Power"),
    SunSpecPoint(19, "VAr_SF", "sunssf", 1, description="Reactive Power Scale Factor"),
    SunSpecPoint(20, "PF", "int16", 1, scale_factor_point="PF_SF", unit="pct", sunpm_tag="PF", description="Power Factor"),
    SunSpecPoint(21, "PF_SF", "sunssf", 1, description="Power Factor Scale Factor"),
    SunSpecPoint(22, "WH", "uint32", 2, scale_factor_point="WH_SF", unit="Wh", sunpm_tag="E_TOT", description="Lifetime Energy Production"),
    SunSpecPoint(24, "WH_SF", "sunssf", 1, description="Energy Scale Factor"),
    SunSpecPoint(25, "DCA", "uint16", 1, scale_factor_point="DCA_SF", unit="A", sunpm_tag="I_DC", description="DC Current Total"),
    SunSpecPoint(26, "DCA_SF", "sunssf", 1, description="DC Current Scale Factor"),
    SunSpecPoint(27, "DCV", "uint16", 1, scale_factor_point="DCV_SF", unit="V", sunpm_tag="U_DC", description="DC Voltage Total"),
    SunSpecPoint(28, "DCV_SF", "sunssf", 1, description="DC Voltage Scale Factor"),
    SunSpecPoint(29, "DCW", "int16", 1, scale_factor_point="DCW_SF", unit="W", sunpm_tag="P_DC", description="DC Power Total"),
    SunSpecPoint(30, "DCW_SF", "sunssf", 1, description="DC Power Scale Factor"),
    SunSpecPoint(31, "TmpCab", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", sunpm_tag="TEMP_INT", description="Cabinet Temperature"),
    SunSpecPoint(32, "TmpSnk", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", sunpm_tag="TEMP_EXT", description="Heat Sink Temperature"),
    SunSpecPoint(33, "TmpTrns", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", description="Transformer Temperature"),
    SunSpecPoint(34, "TmpOt", "int16", 1, scale_factor_point="Tmp_SF", unit="degC", description="Other Temperature"),
    SunSpecPoint(35, "Tmp_SF", "sunssf", 1, description="Temperature Scale Factor"),
    SunSpecPoint(36, "St", "enum16", 1, sunpm_tag="STATE", description="Inverter Operating State"),
    SunSpecPoint(37, "StVnd", "enum16", 1, description="Vendor Specific Operating State"),
    SunSpecPoint(38, "Evt1", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="Inverter Event 1 / Alarms"),
    SunSpecPoint(40, "Evt2", "bitfield32", 2, action_code="1", description="Inverter Event 2 / Warnings"),
    SunSpecPoint(42, "EvtVnd1", "bitfield32", 2, action_code="1", description="Vendor Event 1"),
    SunSpecPoint(44, "EvtVnd2", "bitfield32", 2, action_code="1", description="Vendor Event 2"),
    SunSpecPoint(46, "EvtVnd3", "bitfield32", 2, action_code="1", description="Vendor Event 3"),
    SunSpecPoint(48, "EvtVnd4", "bitfield32", 2, action_code="1", description="Vendor Event 4"),
]

# ==============================================================================
# MODEL 120 : Nameplate Ratings
# ==============================================================================
MODEL_120_POINTS = [
    SunSpecPoint(0, "DERTyp", "enum16", 1, description="DER Device Type"),
    SunSpecPoint(1, "WRtg", "uint16", 1, scale_factor_point="WRtg_SF", unit="W", description="Continuous Active Power Rating"),
    SunSpecPoint(2, "WRtg_SF", "sunssf", 1, description="Active Power Rating Scale Factor"),
    SunSpecPoint(3, "VARtg", "uint16", 1, scale_factor_point="VARtg_SF", unit="VA", description="Continuous Apparent Power Rating"),
    SunSpecPoint(4, "VARtg_SF", "sunssf", 1, description="Apparent Power Rating Scale Factor"),
    SunSpecPoint(5, "VArRtgQ1", "int16", 1, scale_factor_point="VArRtg_SF", unit="var", description="Continuous Reactive Power Q1"),
    SunSpecPoint(6, "VArRtgQ2", "int16", 1, scale_factor_point="VArRtg_SF", unit="var", description="Continuous Reactive Power Q2"),
    SunSpecPoint(7, "VArRtgQ3", "int16", 1, scale_factor_point="VArRtg_SF", unit="var", description="Continuous Reactive Power Q3"),
    SunSpecPoint(8, "VArRtgQ4", "int16", 1, scale_factor_point="VArRtg_SF", unit="var", description="Continuous Reactive Power Q4"),
    SunSpecPoint(9, "VArRtg_SF", "sunssf", 1, description="Reactive Power Rating Scale Factor"),
    SunSpecPoint(10, "ARtg", "uint16", 1, scale_factor_point="ARtg_SF", unit="A", description="Continuous AC Current Rating"),
    SunSpecPoint(11, "ARtg_SF", "sunssf", 1, description="Current Rating Scale Factor"),
    SunSpecPoint(12, "PFRtgQ1", "int16", 1, scale_factor_point="PFRtg_SF", unit="cosPhi", description="Minimum Power Factor Q1"),
    SunSpecPoint(13, "PFRtgQ2", "int16", 1, scale_factor_point="PFRtg_SF", unit="cosPhi", description="Minimum Power Factor Q2"),
    SunSpecPoint(14, "PFRtgQ3", "int16", 1, scale_factor_point="PFRtg_SF", unit="cosPhi", description="Minimum Power Factor Q3"),
    SunSpecPoint(15, "PFRtgQ4", "int16", 1, scale_factor_point="PFRtg_SF", unit="cosPhi", description="Minimum Power Factor Q4"),
    SunSpecPoint(16, "PFRtg_SF", "sunssf", 1, description="Power Factor Rating Scale Factor"),
    SunSpecPoint(17, "WHRtg", "uint16", 1, scale_factor_point="WHRtg_SF", unit="Wh", description="Nominal Energy Storage Capacity"),
    SunSpecPoint(18, "WHRtg_SF", "sunssf", 1, description="Energy Capacity Scale Factor"),
    SunSpecPoint(19, "AhrRtg", "uint16", 1, scale_factor_point="AhrRtg_SF", unit="Ah", description="Usable Amp-Hour Capacity"),
    SunSpecPoint(20, "AhrRtg_SF", "sunssf", 1, description="Amp-Hour Capacity Scale Factor"),
    SunSpecPoint(21, "MaxChaRte", "uint16", 1, scale_factor_point="MaxChaRte_SF", unit="W", description="Maximum Battery Charge Rate"),
    SunSpecPoint(22, "MaxChaRte_SF", "sunssf", 1, description="Max Charge Rate Scale Factor"),
    SunSpecPoint(23, "MaxDisChaRte", "uint16", 1, scale_factor_point="MaxDisChaRte_SF", unit="W", description="Maximum Battery Discharge Rate"),
    SunSpecPoint(24, "MaxDisChaRte_SF", "sunssf", 1, description="Max Discharge Rate Scale Factor"),
    SunSpecPoint(25, "Pad", "pad", 1, description="Alignment Pad"),
]

# ==============================================================================
# MODEL 123 : Immediate Inverter Controls (Power Curtailment & Power Factor)
# ==============================================================================
MODEL_123_POINTS = [
    SunSpecPoint(0, "Conn_WinTms", "uint16", 1, unit="s", access="RW", action_code="8", description="Connection Randomized Time Window"),
    SunSpecPoint(1, "Conn_RvrtTms", "uint16", 1, unit="s", access="RW", action_code="8", description="Connection Revert Time"),
    SunSpecPoint(2, "Conn", "enum16", 1, access="RW", action_code="8", description="Inverter Connection Control Switch (0: Disconnect, 1: Connect)"),
    SunSpecPoint(3, "WMaxLimPct", "uint16", 1, scale_factor_point="WMaxLimPct_SF", unit="pct", access="RW", action_code="8", description="Active Power Curtailment Limit (% of rated W)"),
    SunSpecPoint(4, "WMaxLimPct_WinTms", "uint16", 1, unit="s", access="RW", action_code="8", description="Power Limit Window Time"),
    SunSpecPoint(5, "WMaxLimPct_RvrtTms", "uint16", 1, unit="s", access="RW", action_code="8", description="Power Limit Revert Time"),
    SunSpecPoint(6, "WMaxLimPct_RmpTms", "uint16", 1, unit="s", access="RW", action_code="8", description="Power Limit Ramp Time"),
    SunSpecPoint(7, "WMaxLimPct_Ena", "enum16", 1, access="RW", action_code="8", description="Enable Active Power Limit Control"),
    SunSpecPoint(8, "OutPFSet", "int16", 1, scale_factor_point="OutPFSet_SF", unit="cosPhi", access="RW", action_code="8", description="Power Factor Setpoint"),
    SunSpecPoint(9, "OutPFSet_WinTms", "uint16", 1, unit="s", access="RW", action_code="8", description="Power Factor Window Time"),
    SunSpecPoint(10, "OutPFSet_RvrtTms", "uint16", 1, unit="s", access="RW", action_code="8", description="Power Factor Revert Time"),
    SunSpecPoint(11, "OutPFSet_RmpTms", "uint16", 1, unit="s", access="RW", action_code="8", description="Power Factor Ramp Time"),
    SunSpecPoint(12, "OutPFSet_Ena", "enum16", 1, access="RW", action_code="8", description="Enable Power Factor Setpoint"),
    SunSpecPoint(13, "VArWMaxPct", "int16", 1, scale_factor_point="VArPct_SF", unit="pct", access="RW", action_code="8", description="Reactive Power Limit % of WMax"),
    SunSpecPoint(14, "VArMaxPct", "int16", 1, scale_factor_point="VArPct_SF", unit="pct", access="RW", action_code="8", description="Reactive Power Limit % of VArMax"),
    SunSpecPoint(15, "VArAvalPct", "int16", 1, scale_factor_point="VArPct_SF", unit="pct", access="RW", action_code="8", description="Reactive Power Limit % of Available VAr"),
    SunSpecPoint(16, "VArPct_Ena", "enum16", 1, access="RW", action_code="8", description="Enable Reactive Power Control"),
    SunSpecPoint(17, "WMaxLim_SF", "sunssf", 1, description="Active Power Limit Scale Factor"),
    SunSpecPoint(18, "OutPFSet_SF", "sunssf", 1, description="Power Factor Setpoint Scale Factor"),
    SunSpecPoint(19, "VArPct_SF", "sunssf", 1, description="Reactive Power Scale Factor"),
]

# ==============================================================================
# MODEL 124 : Storage Control (BESS in Hybrid Inverters)
# ==============================================================================
MODEL_124_POINTS = [
    SunSpecPoint(0, "WChaMax", "uint16", 1, scale_factor_point="WChaMax_SF", unit="W", access="RW", action_code="8", description="Storage Set Max Charge Rate"),
    SunSpecPoint(1, "WChaGra", "uint16", 1, scale_factor_point="WChaDisChaGra_SF", unit="pctWMax/s", access="RW", action_code="8", description="Storage Charge Ramp Rate"),
    SunSpecPoint(2, "WDisChaGra", "uint16", 1, scale_factor_point="WChaDisChaGra_SF", unit="pctWMax/s", access="RW", action_code="8", description="Storage Discharge Ramp Rate"),
    SunSpecPoint(3, "StorCtl_Mod", "bitfield16", 1, access="RW", action_code="8", description="Storage Control Mode"),
    SunSpecPoint(4, "VAChaMax", "uint16", 1, scale_factor_point="VAChaMax_SF", unit="VA", access="RW", action_code="8", description="Storage Set Max Charge VA"),
    SunSpecPoint(5, "MinRsvPct", "uint16", 1, scale_factor_point="MinRsvPct_SF", unit="pct", access="RW", action_code="8", description="Storage Minimum Reserve SoC %"),
    SunSpecPoint(6, "ChaState", "uint16", 1, scale_factor_point="ChaState_SF", unit="pct", sunpm_tag="SOC", description="Battery State of Charge (SoC)"),
    SunSpecPoint(7, "StorAval", "uint16", 1, scale_factor_point="StorAval_SF", unit="Ah", description="Storage Available Energy"),
    SunSpecPoint(8, "InBatV", "uint16", 1, scale_factor_point="InBatV_SF", unit="V", sunpm_tag="U_DC", description="Internal Battery DC Voltage"),
    SunSpecPoint(9, "ChaSt", "enum16", 1, description="Charge / Discharge Status"),
    SunSpecPoint(10, "OutBatV", "uint16", 1, scale_factor_point="OutBatV_SF", unit="V", description="External Battery Terminal Voltage"),
    SunSpecPoint(11, "WChaMax_SF", "sunssf", 1, description="Max Charge Rate Scale Factor"),
    SunSpecPoint(12, "WChaDisChaGra_SF", "sunssf", 1, description="Ramp Rate Scale Factor"),
    SunSpecPoint(13, "VAChaMax_SF", "sunssf", 1, description="Charge VA Scale Factor"),
    SunSpecPoint(14, "MinRsvPct_SF", "sunssf", 1, description="Reserve Scale Factor"),
    SunSpecPoint(15, "ChaState_SF", "sunssf", 1, description="State of Charge Scale Factor"),
    SunSpecPoint(16, "StorAval_SF", "sunssf", 1, description="Available Storage Scale Factor"),
    SunSpecPoint(17, "InBatV_SF", "sunssf", 1, description="Battery Voltage Scale Factor"),
    SunSpecPoint(18, "OutBatV_SF", "sunssf", 1, description="Terminal Voltage Scale Factor"),
]

# ==============================================================================
# MODEL 160 : Multiple MPPT Inverter Extension Model
# ==============================================================================
MODEL_160_HEADER_POINTS = [
    SunSpecPoint(0, "DCA_SF", "sunssf", 1, description="MPPT Current Scale Factor"),
    SunSpecPoint(1, "DCV_SF", "sunssf", 1, description="MPPT Voltage Scale Factor"),
    SunSpecPoint(2, "DCW_SF", "sunssf", 1, description="MPPT Power Scale Factor"),
    SunSpecPoint(3, "DCWH_SF", "sunssf", 1, description="MPPT Energy Scale Factor"),
    SunSpecPoint(4, "Evt", "bitfield32", 2, action_code="1", description="Global MPPT Events"),
    SunSpecPoint(6, "N", "uint16", 1, description="Number of MPPT Modules / Inputs"),
    SunSpecPoint(7, "TmsPer", "uint16", 1, unit="s", description="Period of Timestamp"),
]

MODEL_160_REPEATING_TEMPLATE = [
    SunSpecPoint(0, "ID", "uint16", 1, description="MPPT Input #{i} ID"),
    SunSpecPoint(1, "IDStr", "string8", 8, description="MPPT Input #{i} String Name"),
    SunSpecPoint(9, "DCA", "uint16", 1, scale_factor_point="DCA_SF", unit="A", sunpm_tag="I_DC_{i}", description="MPPT #{i} Current"),
    SunSpecPoint(10, "DCV", "uint16", 1, scale_factor_point="DCV_SF", unit="V", sunpm_tag="U_DC_{i}", description="MPPT #{i} Voltage"),
    SunSpecPoint(11, "DCW", "uint16", 1, scale_factor_point="DCW_SF", unit="W", sunpm_tag="P_DC_{i}", description="MPPT #{i} Power"),
    SunSpecPoint(12, "DCWH", "uint32", 2, scale_factor_point="DCWH_SF", unit="Wh", description="MPPT #{i} Lifetime Energy"),
    SunSpecPoint(14, "Tms", "uint32", 2, unit="s", description="MPPT #{i} Timestamp"),
    SunSpecPoint(16, "Tmp", "int16", 1, unit="degC", description="MPPT #{i} Operating Temperature"),
    SunSpecPoint(17, "DCSt", "enum16", 1, description="MPPT #{i} Operating State"),
    SunSpecPoint(18, "DCEvt", "bitfield32", 2, action_code="1", description="MPPT #{i} Events"),
]

# ==============================================================================
# MODEL 203 : Three Phase Wye Meter
# ==============================================================================
MODEL_203_POINTS = [
    SunSpecPoint(0, "A", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC", description="Total AC Current"),
    SunSpecPoint(1, "AphA", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_1", description="Phase A Current"),
    SunSpecPoint(2, "AphB", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_2", description="Phase B Current"),
    SunSpecPoint(3, "AphC", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_3", description="Phase C Current"),
    SunSpecPoint(4, "A_SF", "sunssf", 1, description="Current Scale Factor"),
    SunSpecPoint(5, "PhV", "int16", 1, scale_factor_point="V_SF", unit="V", description="Average Phase-to-Neutral Voltage"),
    SunSpecPoint(6, "PhVphA", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_1", description="Phase A Voltage"),
    SunSpecPoint(7, "PhVphB", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_2", description="Phase B Voltage"),
    SunSpecPoint(8, "PhVphC", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_3", description="Phase C Voltage"),
    SunSpecPoint(9, "PPV", "int16", 1, scale_factor_point="V_SF", unit="V", description="Average Phase-to-Phase Voltage"),
    SunSpecPoint(10, "PPVphAB", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L1L2", description="Phase Voltage AB"),
    SunSpecPoint(11, "PPVphBC", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L2L3", description="Phase Voltage BC"),
    SunSpecPoint(12, "PPVphCA", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L3L1", description="Phase Voltage CA"),
    SunSpecPoint(13, "V_SF", "sunssf", 1, description="Voltage Scale Factor"),
    SunSpecPoint(14, "Hz", "int16", 1, scale_factor_point="Hz_SF", unit="Hz", sunpm_tag="FREQ", description="Frequency"),
    SunSpecPoint(15, "Hz_SF", "sunssf", 1, description="Frequency Scale Factor"),
    SunSpecPoint(16, "W", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC", description="Total Real Power"),
    SunSpecPoint(17, "WphA", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_1", description="Phase A Real Power"),
    SunSpecPoint(18, "WphB", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_2", description="Phase B Real Power"),
    SunSpecPoint(19, "WphC", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_3", description="Phase C Real Power"),
    SunSpecPoint(20, "W_SF", "sunssf", 1, description="Real Power Scale Factor"),
    SunSpecPoint(21, "VA", "int16", 1, scale_factor_point="VA_SF", unit="VA", sunpm_tag="S_AC", description="Total Apparent Power"),
    SunSpecPoint(22, "VAphA", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase A Apparent Power"),
    SunSpecPoint(23, "VAphB", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase B Apparent Power"),
    SunSpecPoint(24, "VAphC", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase C Apparent Power"),
    SunSpecPoint(25, "VA_SF", "sunssf", 1, description="Apparent Power Scale Factor"),
    SunSpecPoint(26, "VAR", "int16", 1, scale_factor_point="VAR_SF", unit="var", sunpm_tag="Q_AC", description="Total Reactive Power"),
    SunSpecPoint(27, "VARphA", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase A Reactive Power"),
    SunSpecPoint(28, "VARphB", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase B Reactive Power"),
    SunSpecPoint(29, "VARphC", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase C Reactive Power"),
    SunSpecPoint(30, "VAR_SF", "sunssf", 1, description="Reactive Power Scale Factor"),
    SunSpecPoint(31, "PF", "int16", 1, scale_factor_point="PF_SF", unit="pct", sunpm_tag="PF", description="Total Power Factor"),
    SunSpecPoint(32, "PFphA", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase A Power Factor"),
    SunSpecPoint(33, "PFphB", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase B Power Factor"),
    SunSpecPoint(34, "PFphC", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase C Power Factor"),
    SunSpecPoint(35, "PF_SF", "sunssf", 1, description="Power Factor Scale Factor"),
    SunSpecPoint(36, "TotWhExp", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", sunpm_tag="E_TOT", description="Total Real Energy Exported"),
    SunSpecPoint(38, "TotWhImp", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", description="Total Real Energy Imported"),
    SunSpecPoint(40, "TotWh_SF", "sunssf", 1, description="Energy Scale Factor"),
    SunSpecPoint(41, "Events", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="Meter Alarms / Events"),
]

# ==============================================================================
# MODEL 201 : Single Phase AC Meter
# ==============================================================================
MODEL_201_POINTS = [
    SunSpecPoint(0, "A", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC", description="Total AC Current"),
    SunSpecPoint(1, "AphA", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_1", description="Phase A Current"),
    SunSpecPoint(2, "A_SF", "sunssf", 1, description="Current Scale Factor"),
    SunSpecPoint(3, "PhV", "int16", 1, scale_factor_point="V_SF", unit="V", description="Average Phase-to-Neutral Voltage"),
    SunSpecPoint(4, "PhVphA", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_1", description="Phase A Voltage"),
    SunSpecPoint(5, "V_SF", "sunssf", 1, description="Voltage Scale Factor"),
    SunSpecPoint(6, "Hz", "int16", 1, scale_factor_point="Hz_SF", unit="Hz", sunpm_tag="FREQ", description="Frequency"),
    SunSpecPoint(7, "Hz_SF", "sunssf", 1, description="Frequency Scale Factor"),
    SunSpecPoint(8, "W", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC", description="Total Real Power"),
    SunSpecPoint(9, "WphA", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_1", description="Phase A Real Power"),
    SunSpecPoint(10, "W_SF", "sunssf", 1, description="Real Power Scale Factor"),
    SunSpecPoint(11, "VA", "int16", 1, scale_factor_point="VA_SF", unit="VA", sunpm_tag="S_AC", description="Total Apparent Power"),
    SunSpecPoint(12, "VAphA", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase A Apparent Power"),
    SunSpecPoint(13, "VA_SF", "sunssf", 1, description="Apparent Power Scale Factor"),
    SunSpecPoint(14, "VAR", "int16", 1, scale_factor_point="VAR_SF", unit="var", sunpm_tag="Q_AC", description="Total Reactive Power"),
    SunSpecPoint(15, "VARphA", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase A Reactive Power"),
    SunSpecPoint(16, "VAR_SF", "sunssf", 1, description="Reactive Power Scale Factor"),
    SunSpecPoint(17, "PF", "int16", 1, scale_factor_point="PF_SF", unit="pct", sunpm_tag="PF", description="Total Power Factor"),
    SunSpecPoint(18, "PFphA", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase A Power Factor"),
    SunSpecPoint(19, "PF_SF", "sunssf", 1, description="Power Factor Scale Factor"),
    SunSpecPoint(20, "TotWhExp", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", sunpm_tag="E_TOT", description="Total Real Energy Exported"),
    SunSpecPoint(22, "TotWhImp", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", description="Total Real Energy Imported"),
    SunSpecPoint(24, "TotWh_SF", "sunssf", 1, description="Energy Scale Factor"),
    SunSpecPoint(25, "Events", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="Meter Alarms / Events"),
]

# ==============================================================================
# MODEL 202 : Split Phase AC Meter
# ==============================================================================
MODEL_202_POINTS = [
    SunSpecPoint(0, "A", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC", description="Total AC Current"),
    SunSpecPoint(1, "AphA", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_1", description="Phase A Current"),
    SunSpecPoint(2, "AphB", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_2", description="Phase B Current"),
    SunSpecPoint(3, "A_SF", "sunssf", 1, description="Current Scale Factor"),
    SunSpecPoint(4, "PhV", "int16", 1, scale_factor_point="V_SF", unit="V", description="Average Phase-to-Neutral Voltage"),
    SunSpecPoint(5, "PhVphA", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_1", description="Phase A Voltage"),
    SunSpecPoint(6, "PhVphB", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_2", description="Phase B Voltage"),
    SunSpecPoint(7, "PPV", "int16", 1, scale_factor_point="V_SF", unit="V", description="Average Phase-to-Phase Voltage"),
    SunSpecPoint(8, "PPVphAB", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L1L2", description="Phase Voltage AB"),
    SunSpecPoint(9, "V_SF", "sunssf", 1, description="Voltage Scale Factor"),
    SunSpecPoint(10, "Hz", "int16", 1, scale_factor_point="Hz_SF", unit="Hz", sunpm_tag="FREQ", description="Frequency"),
    SunSpecPoint(11, "Hz_SF", "sunssf", 1, description="Frequency Scale Factor"),
    SunSpecPoint(12, "W", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC", description="Total Real Power"),
    SunSpecPoint(13, "WphA", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_1", description="Phase A Real Power"),
    SunSpecPoint(14, "WphB", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_2", description="Phase B Real Power"),
    SunSpecPoint(15, "W_SF", "sunssf", 1, description="Real Power Scale Factor"),
    SunSpecPoint(16, "VA", "int16", 1, scale_factor_point="VA_SF", unit="VA", sunpm_tag="S_AC", description="Total Apparent Power"),
    SunSpecPoint(17, "VAphA", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase A Apparent Power"),
    SunSpecPoint(18, "VAphB", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase B Apparent Power"),
    SunSpecPoint(19, "VA_SF", "sunssf", 1, description="Apparent Power Scale Factor"),
    SunSpecPoint(20, "VAR", "int16", 1, scale_factor_point="VAR_SF", unit="var", sunpm_tag="Q_AC", description="Total Reactive Power"),
    SunSpecPoint(21, "VARphA", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase A Reactive Power"),
    SunSpecPoint(22, "VARphB", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase B Reactive Power"),
    SunSpecPoint(23, "VAR_SF", "sunssf", 1, description="Reactive Power Scale Factor"),
    SunSpecPoint(24, "PF", "int16", 1, scale_factor_point="PF_SF", unit="pct", sunpm_tag="PF", description="Total Power Factor"),
    SunSpecPoint(25, "PFphA", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase A Power Factor"),
    SunSpecPoint(26, "PFphB", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase B Power Factor"),
    SunSpecPoint(27, "PF_SF", "sunssf", 1, description="Power Factor Scale Factor"),
    SunSpecPoint(28, "TotWhExp", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", sunpm_tag="E_TOT", description="Total Real Energy Exported"),
    SunSpecPoint(30, "TotWhImp", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", description="Total Real Energy Imported"),
    SunSpecPoint(32, "TotWh_SF", "sunssf", 1, description="Energy Scale Factor"),
    SunSpecPoint(33, "Events", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="Meter Alarms / Events"),
]

# ==============================================================================
# MODEL 204 : Three Phase Delta AC Meter
# ==============================================================================
MODEL_204_POINTS = [
    SunSpecPoint(0, "A", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC", description="Total AC Current"),
    SunSpecPoint(1, "AphA", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_1", description="Phase A Current"),
    SunSpecPoint(2, "AphB", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_2", description="Phase B Current"),
    SunSpecPoint(3, "AphC", "int16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_3", description="Phase C Current"),
    SunSpecPoint(4, "A_SF", "sunssf", 1, description="Current Scale Factor"),
    SunSpecPoint(5, "PPV", "int16", 1, scale_factor_point="V_SF", unit="V", description="Average Phase-to-Phase Voltage"),
    SunSpecPoint(6, "PPVphAB", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L1L2", description="Phase Voltage AB"),
    SunSpecPoint(7, "PPVphBC", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L2L3", description="Phase Voltage BC"),
    SunSpecPoint(8, "PPVphCA", "int16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L3L1", description="Phase Voltage CA"),
    SunSpecPoint(9, "V_SF", "sunssf", 1, description="Voltage Scale Factor"),
    SunSpecPoint(10, "Hz", "int16", 1, scale_factor_point="Hz_SF", unit="Hz", sunpm_tag="FREQ", description="Frequency"),
    SunSpecPoint(11, "Hz_SF", "sunssf", 1, description="Frequency Scale Factor"),
    SunSpecPoint(12, "W", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC", description="Total Real Power"),
    SunSpecPoint(13, "WphA", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_1", description="Phase A Real Power"),
    SunSpecPoint(14, "WphB", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_2", description="Phase B Real Power"),
    SunSpecPoint(15, "WphC", "int16", 1, scale_factor_point="W_SF", unit="W", sunpm_tag="P_AC_3", description="Phase C Real Power"),
    SunSpecPoint(16, "W_SF", "sunssf", 1, description="Real Power Scale Factor"),
    SunSpecPoint(17, "VA", "int16", 1, scale_factor_point="VA_SF", unit="VA", sunpm_tag="S_AC", description="Total Apparent Power"),
    SunSpecPoint(18, "VAphA", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase A Apparent Power"),
    SunSpecPoint(19, "VAphB", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase B Apparent Power"),
    SunSpecPoint(20, "VAphC", "int16", 1, scale_factor_point="VA_SF", unit="VA", description="Phase C Apparent Power"),
    SunSpecPoint(21, "VA_SF", "sunssf", 1, description="Apparent Power Scale Factor"),
    SunSpecPoint(22, "VAR", "int16", 1, scale_factor_point="VAR_SF", unit="var", sunpm_tag="Q_AC", description="Total Reactive Power"),
    SunSpecPoint(23, "VARphA", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase A Reactive Power"),
    SunSpecPoint(24, "VARphB", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase B Reactive Power"),
    SunSpecPoint(25, "VARphC", "int16", 1, scale_factor_point="VAR_SF", unit="var", description="Phase C Reactive Power"),
    SunSpecPoint(26, "VAR_SF", "sunssf", 1, description="Reactive Power Scale Factor"),
    SunSpecPoint(27, "PF", "int16", 1, scale_factor_point="PF_SF", unit="pct", sunpm_tag="PF", description="Total Power Factor"),
    SunSpecPoint(28, "PFphA", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase A Power Factor"),
    SunSpecPoint(29, "PFphB", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase B Power Factor"),
    SunSpecPoint(30, "PFphC", "int16", 1, scale_factor_point="PF_SF", unit="cosPhi", description="Phase C Power Factor"),
    SunSpecPoint(31, "PF_SF", "sunssf", 1, description="Power Factor Scale Factor"),
    SunSpecPoint(32, "TotWhExp", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", sunpm_tag="E_TOT", description="Total Real Energy Exported"),
    SunSpecPoint(34, "TotWhImp", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", description="Total Real Energy Imported"),
    SunSpecPoint(36, "TotWh_SF", "sunssf", 1, description="Energy Scale Factor"),
    SunSpecPoint(37, "Events", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="Meter Alarms / Events"),
]

# ==============================================================================
# MODEL 701 : DER AC Measurement (Modern IEEE 1547-2018 DER AC Point Map)
# ==============================================================================
MODEL_701_POINTS = [
    SunSpecPoint(0, "A_A", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_1", description="DER Phase A Current"),
    SunSpecPoint(1, "A_B", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_2", description="DER Phase B Current"),
    SunSpecPoint(2, "A_C", "uint16", 1, scale_factor_point="A_SF", unit="A", sunpm_tag="I_AC_3", description="DER Phase C Current"),
    SunSpecPoint(3, "A_SF", "sunssf", 1, description="DER Current Scale Factor"),
    SunSpecPoint(4, "V_A", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_1", description="DER Phase A Voltage"),
    SunSpecPoint(5, "V_B", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_2", description="DER Phase B Voltage"),
    SunSpecPoint(6, "V_C", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_3", description="DER Phase C Voltage"),
    SunSpecPoint(7, "V_AB", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L1L2", description="DER Voltage Phase AB"),
    SunSpecPoint(8, "V_BC", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L2L3", description="DER Voltage Phase BC"),
    SunSpecPoint(9, "V_CA", "uint16", 1, scale_factor_point="V_SF", unit="V", sunpm_tag="U_AC_L3L1", description="DER Voltage Phase CA"),
    SunSpecPoint(10, "V_SF", "sunssf", 1, description="DER Voltage Scale Factor"),
    SunSpecPoint(11, "Hz", "uint16", 1, scale_factor_point="Hz_SF", unit="Hz", sunpm_tag="FREQ", description="DER AC Frequency"),
    SunSpecPoint(12, "Hz_SF", "sunssf", 1, description="DER Frequency Scale Factor"),
    SunSpecPoint(13, "TotW", "int16", 1, scale_factor_point="TotW_SF", unit="W", sunpm_tag="P_AC", description="DER Total Active Power"),
    SunSpecPoint(14, "W_A", "int16", 1, scale_factor_point="TotW_SF", unit="W", sunpm_tag="P_AC_1", description="DER Phase A Active Power"),
    SunSpecPoint(15, "W_B", "int16", 1, scale_factor_point="TotW_SF", unit="W", sunpm_tag="P_AC_2", description="DER Phase B Active Power"),
    SunSpecPoint(16, "W_C", "int16", 1, scale_factor_point="TotW_SF", unit="W", sunpm_tag="P_AC_3", description="DER Phase C Active Power"),
    SunSpecPoint(17, "TotW_SF", "sunssf", 1, description="DER Power Scale Factor"),
    SunSpecPoint(18, "TotVA", "int16", 1, scale_factor_point="TotVA_SF", unit="VA", sunpm_tag="S_AC", description="DER Total Apparent Power"),
    SunSpecPoint(19, "TotVA_SF", "sunssf", 1, description="DER Apparent Power Scale Factor"),
    SunSpecPoint(20, "TotVAr", "int16", 1, scale_factor_point="TotVAr_SF", unit="var", sunpm_tag="Q_AC", description="DER Total Reactive Power"),
    SunSpecPoint(21, "TotVAr_SF", "sunssf", 1, description="DER Reactive Power Scale Factor"),
    SunSpecPoint(22, "TotPF", "int16", 1, scale_factor_point="TotPF_SF", unit="pct", sunpm_tag="PF", description="DER Power Factor"),
    SunSpecPoint(23, "TotPF_SF", "sunssf", 1, description="DER Power Factor Scale Factor"),
    SunSpecPoint(24, "TotWh", "uint32", 2, scale_factor_point="TotWh_SF", unit="Wh", sunpm_tag="E_TOT", description="DER Total Energy Production"),
    SunSpecPoint(26, "TotWh_SF", "sunssf", 1, description="DER Energy Scale Factor"),
    SunSpecPoint(27, "DERSt", "enum16", 1, sunpm_tag="STATE", description="DER Operational Status"),
    SunSpecPoint(28, "Alrm", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="DER Active Alarms"),
]

# ==============================================================================
# MODEL 702 : DER Capacity / Nameplate
# ==============================================================================
MODEL_702_POINTS = [
    SunSpecPoint(0, "ActivePRtg", "uint16", 1, scale_factor_point="ActivePRtg_SF", unit="W", description="DER Continuous Active Power Rating"),
    SunSpecPoint(1, "ActivePRtg_SF", "sunssf", 1, description="Active Power Rating Scale Factor"),
    SunSpecPoint(2, "AppPRtg", "uint16", 1, scale_factor_point="AppPRtg_SF", unit="VA", description="DER Apparent Power Rating"),
    SunSpecPoint(3, "AppPRtg_SF", "sunssf", 1, description="Apparent Power Rating Scale Factor"),
    SunSpecPoint(4, "ReactPRtgQ1", "int16", 1, scale_factor_point="ReactPRtg_SF", unit="var", description="DER Reactive Power Q1"),
    SunSpecPoint(5, "ReactPRtgQ4", "int16", 1, scale_factor_point="ReactPRtg_SF", unit="var", description="DER Reactive Power Q4"),
    SunSpecPoint(6, "ReactPRtg_SF", "sunssf", 1, description="Reactive Power Rating Scale Factor"),
    SunSpecPoint(7, "ARtg", "uint16", 1, scale_factor_point="ARtg_SF", unit="A", description="DER Max AC Current Rating"),
    SunSpecPoint(8, "ARtg_SF", "sunssf", 1, description="AC Current Rating Scale Factor"),
]

# ==============================================================================
# MODEL 704 : DER Active and Reactive Power Controls / Curtailment
# ==============================================================================
MODEL_704_POINTS = [
    SunSpecPoint(0, "WMaxLimPct", "uint16", 1, scale_factor_point="WMaxLimPct_SF", unit="pct", access="RW", action_code="8", description="DER Active Power Curtailment Limit %"),
    SunSpecPoint(1, "WMaxLimPct_Ena", "enum16", 1, access="RW", action_code="8", description="Enable DER Active Power Curtailment"),
    SunSpecPoint(2, "WMaxLimPct_SF", "sunssf", 1, description="Active Power Curtailment Scale Factor"),
    SunSpecPoint(3, "OutPFSet", "int16", 1, scale_factor_point="OutPFSet_SF", unit="cosPhi", access="RW", action_code="8", description="DER Target Power Factor Setpoint"),
    SunSpecPoint(4, "OutPFSet_Ena", "enum16", 1, access="RW", action_code="8", description="Enable DER Power Factor Setpoint"),
    SunSpecPoint(5, "OutPFSet_SF", "sunssf", 1, description="Power Factor Setpoint Scale Factor"),
    SunSpecPoint(6, "VArMaxPct", "int16", 1, scale_factor_point="VArMaxPct_SF", unit="pct", access="RW", action_code="8", description="DER Reactive Power Limit % of Max VAr"),
    SunSpecPoint(7, "VArMaxPct_Ena", "enum16", 1, access="RW", action_code="8", description="Enable DER Reactive Power Limit"),
    SunSpecPoint(8, "VArMaxPct_SF", "sunssf", 1, description="Reactive Power Limit Scale Factor"),
]

# ==============================================================================
# MODEL 714 : DER Multiple DC Ports (Repeating Block per Port)
# ==============================================================================
MODEL_714_HEADER_POINTS = [
    SunSpecPoint(0, "DCA_SF", "sunssf", 1, description="DER DC Port Current Scale Factor"),
    SunSpecPoint(1, "DCV_SF", "sunssf", 1, description="DER DC Port Voltage Scale Factor"),
    SunSpecPoint(2, "DCW_SF", "sunssf", 1, description="DER DC Port Power Scale Factor"),
    SunSpecPoint(3, "DCWh_SF", "sunssf", 1, description="DER DC Port Energy Scale Factor"),
    SunSpecPoint(4, "DCEvt", "bitfield32", 2, action_code="1", description="DER DC Port Global Events"),
    SunSpecPoint(6, "NPrt", "uint16", 1, description="Number of DC Ports (NPrt)"),
    SunSpecPoint(7, "Pad", "pad", 1, description="Header Alignment Pad"),
]

MODEL_714_REPEATING_TEMPLATE = [
    SunSpecPoint(0, "PrtID", "uint16", 1, description="DC Port #{i} ID"),
    SunSpecPoint(1, "PrtTyp", "enum16", 1, description="DC Port #{i} Type (1: PV, 2: BESS)"),
    SunSpecPoint(2, "DCA", "int16", 1, scale_factor_point="DCA_SF", unit="A", sunpm_tag="I_DC_{i}", description="DC Port #{i} Current"),
    SunSpecPoint(3, "DCV", "uint16", 1, scale_factor_point="DCV_SF", unit="V", sunpm_tag="U_DC_{i}", description="DC Port #{i} Voltage"),
    SunSpecPoint(4, "DCW", "int16", 1, scale_factor_point="DCW_SF", unit="W", sunpm_tag="P_DC_{i}", description="DC Port #{i} Power"),
    SunSpecPoint(5, "TotDCWh", "uint32", 2, scale_factor_point="DCWh_SF", unit="Wh", description="DC Port #{i} Total Lifetime Energy"),
    SunSpecPoint(7, "PrtSt", "enum16", 1, description="DC Port #{i} Status"),
    SunSpecPoint(8, "PrtAlrm", "bitfield32", 2, action_code="1", description="DC Port #{i} Alarms"),
    SunSpecPoint(10, "Pad", "pad", 6, description="DC Port #{i} Alignment Pad"),
]

# ==============================================================================
# MODEL 802 : Battery Base Model (BESS Storage)
# ==============================================================================
MODEL_802_POINTS = [
    SunSpecPoint(0, "AhrRtg", "uint16", 1, scale_factor_point="AhrRtg_SF", unit="Ah", description="Nameplate Battery Amp-Hour Capacity"),
    SunSpecPoint(1, "AhrRtg_SF", "sunssf", 1, description="Amp-Hour Capacity Scale Factor"),
    SunSpecPoint(2, "WHRtg", "uint16", 1, scale_factor_point="WHRtg_SF", unit="Wh", description="Nameplate Energy Storage Capacity"),
    SunSpecPoint(3, "WHRtg_SF", "sunssf", 1, description="Energy Capacity Scale Factor"),
    SunSpecPoint(4, "SoC", "uint16", 1, scale_factor_point="SoC_SF", unit="pct", sunpm_tag="SOC", description="State of Charge (SoC)"),
    SunSpecPoint(5, "SoC_SF", "sunssf", 1, description="State of Charge Scale Factor"),
    SunSpecPoint(6, "SoH", "uint16", 1, scale_factor_point="SoH_SF", unit="pct", sunpm_tag="SOH", description="State of Health (SoH)"),
    SunSpecPoint(7, "SoH_SF", "sunssf", 1, description="State of Health Scale Factor"),
    SunSpecPoint(8, "Volts", "uint16", 1, scale_factor_point="Volts_SF", unit="V", sunpm_tag="U_DC", description="Total Battery DC Voltage"),
    SunSpecPoint(9, "Volts_SF", "sunssf", 1, description="Battery Voltage Scale Factor"),
    SunSpecPoint(10, "Amps", "int16", 1, scale_factor_point="Amps_SF", unit="A", sunpm_tag="I_DC", description="Total Battery DC Current"),
    SunSpecPoint(11, "Amps_SF", "sunssf", 1, description="Battery Current Scale Factor"),
    SunSpecPoint(12, "Watts", "int16", 1, scale_factor_point="Watts_SF", unit="W", sunpm_tag="P_DC", description="Total Battery DC Power"),
    SunSpecPoint(13, "Watts_SF", "sunssf", 1, description="Battery Power Scale Factor"),
    SunSpecPoint(14, "State", "enum16", 1, sunpm_tag="STATE", description="Battery Operating State"),
    SunSpecPoint(15, "Alrm", "bitfield32", 2, action_code="1", sunpm_tag="ALARM", description="Battery System Alarms"),
]


# ==============================================================================
# MASTER REGISTRY DICTIONARY
# ==============================================================================
SUNSPEC_MODELS_REGISTRY: dict[int, SunSpecModelDef] = {
    1: SunSpecModelDef(
        model_id=1,
        name="Common",
        description="SunSpec Common Model: Basic device identification & serial",
        standard_length=65,
        points=MODEL_1_POINTS,
    ),
    101: SunSpecModelDef(
        model_id=101,
        name="Inverter_SinglePhase",
        description="Single Phase Inverter Measurements & Status",
        standard_length=50,
        points=MODEL_101_POINTS,
    ),
    102: SunSpecModelDef(
        model_id=102,
        name="Inverter_SplitPhase",
        description="Split Phase Inverter Measurements & Status",
        standard_length=50,
        points=MODEL_102_POINTS,
    ),
    103: SunSpecModelDef(
        model_id=103,
        name="Inverter_ThreePhase",
        description="Three Phase Inverter Measurements & Status",
        standard_length=50,
        points=MODEL_103_POINTS,
    ),
    120: SunSpecModelDef(
        model_id=120,
        name="Nameplate",
        description="Nameplate Ratings & Operating Limits",
        standard_length=26,
        points=MODEL_120_POINTS,
    ),
    123: SunSpecModelDef(
        model_id=123,
        name="ImmediateControls",
        description="Immediate Inverter Controls (Power Curtailment & Power Factor)",
        standard_length=24,
        points=MODEL_123_POINTS,
    ),
    124: SunSpecModelDef(
        model_id=124,
        name="StorageControl",
        description="Storage Controls (BESS Hybrid Inverter Integration)",
        standard_length=24,
        points=MODEL_124_POINTS,
    ),
    160: SunSpecModelDef(
        model_id=160,
        name="MultipleMPPT",
        description="Multiple MPPT Inverter Extension Model",
        standard_length=8,
        points=MODEL_160_HEADER_POINTS,
        is_variable_length=True,
        repeating_block_length=20,
        repeating_points_template=MODEL_160_REPEATING_TEMPLATE,
    ),
    201: SunSpecModelDef(
        model_id=201,
        name="Meter_SinglePhase",
        description="Single Phase AC Energy Meter",
        standard_length=105,
        points=MODEL_201_POINTS,
    ),
    202: SunSpecModelDef(
        model_id=202,
        name="Meter_SplitPhase",
        description="Split Phase AC Energy Meter",
        standard_length=105,
        points=MODEL_202_POINTS,
    ),
    203: SunSpecModelDef(
        model_id=203,
        name="Meter_ThreePhaseWye",
        description="Three Phase (Wye) AC Energy Meter",
        standard_length=105,
        points=MODEL_203_POINTS,
    ),
    204: SunSpecModelDef(
        model_id=204,
        name="Meter_ThreePhaseDelta",
        description="Three Phase (Delta) AC Energy Meter",
        standard_length=105,
        points=MODEL_204_POINTS,
    ),
    701: SunSpecModelDef(
        model_id=701,
        name="DER_AC_Measurement",
        description="DER Modern AC Measurement Model (IEEE 1547-2018)",
        standard_length=153,
        points=MODEL_701_POINTS,
    ),
    702: SunSpecModelDef(
        model_id=702,
        name="DER_Capacity",
        description="DER Capacity & Nameplate Ratings",
        standard_length=38,
        points=MODEL_702_POINTS,
    ),
    704: SunSpecModelDef(
        model_id=704,
        name="DER_Controls",
        description="DER Active and Reactive Power Controls",
        standard_length=108,
        points=MODEL_704_POINTS,
    ),
    714: SunSpecModelDef(
        model_id=714,
        name="DER_MultipleDCPorts",
        description="DER Multiple DC Ports Model",
        standard_length=8,
        points=MODEL_714_HEADER_POINTS,
        is_variable_length=True,
        repeating_block_length=16,
        repeating_points_template=MODEL_714_REPEATING_TEMPLATE,
    ),
    802: SunSpecModelDef(
        model_id=802,
        name="Battery_Base",
        description="Battery Energy Storage System (BESS) Base Model",
        standard_length=64,
        points=MODEL_802_POINTS,
    ),
}


def get_sunspec_model(model_id: int) -> SunSpecModelDef | None:
    """Returns official model definition or None if unknown."""
    return SUNSPEC_MODELS_REGISTRY.get(model_id)


def list_supported_models() -> list[int]:
    """Returns sorted list of all supported official SunSpec Model IDs."""
    return sorted(SUNSPEC_MODELS_REGISTRY.keys())


# Alias for backwards compatibility
SUNSPEC_MODELS: dict[int, SunSpecModelDef] = SUNSPEC_MODELS_REGISTRY
