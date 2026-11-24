"""
SunSpec Alliance Protocol Constants and Standards Reference.

Defines protocol identifiers, standard base addresses, model categories,
sentinel null-values, and generation modes for Webdyn definition files.
"""

from __future__ import annotations

from enum import Enum

# SunSpec Magic Marker: 'SunS' = 0x5375, 0x6E53 (Big-Endian uint32 0x53756E53 = 1399748179)
SUNS_MAGIC_STR: str = "SunS"
SUNS_MAGIC_HEX: int = 0x53756E53
SUNS_MAGIC_HIGH_WORD: int = 0x5375  # 21365
SUNS_MAGIC_LOW_WORD: int = 0x6E53   # 28243

# End of SunSpec model chain sentinel
SUNS_END_MODEL_ID: int = 0xFFFF  # 65535

# Standard Modbus Base Addresses for SunSpec
BASE_ADDRESS_0: int = 40000  # 0-based holding register protocol address
BASE_ADDRESS_1: int = 40001  # 1-based holding register protocol address
BASE_ADDRESS_30000: int = 30000  # Input register base 0 (legacy or alternative)

# Common Model 1 length variants
COMMON_MODEL_LEN_STANDARD: int = 65  # Length without pad register
COMMON_MODEL_LEN_ALIGNED: int = 66   # Length with trailing pad register (64-bit alignment)

# Sentinel Null / Unimplemented Values
SENTINEL_VALUES: dict[str, int | float] = {
    "uint16": 0xFFFF,
    "int16": -32768,  # 0x8000
    "uint32": 0xFFFFFFFF,
    "int32": -2147483648,  # 0x80000000
    "uint64": 0xFFFFFFFFFFFFFFFF,
    "int64": -9223372036854775808,
    "sunssf": -32768,  # 0x8000
    "enum16": 0xFFFF,
    "bitfield16": 0xFFFF,
    "bitfield32": 0xFFFFFFFF,
    "float32": float("nan"),
}

# SunSpec Functional Model Groupings
COMMON_MODELS: frozenset[int] = frozenset({1})

LEGACY_INVERTER_MODELS: frozenset[int] = frozenset({101, 102, 103})
LEGACY_INVERTER_SINGLE_PHASE: frozenset[int] = frozenset({101})
LEGACY_INVERTER_SPLIT_PHASE: frozenset[int] = frozenset({102})
LEGACY_INVERTER_THREE_PHASE: frozenset[int] = frozenset({103})

INVERTER_CONTROL_MODELS: frozenset[int] = frozenset({
    120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132
})

MPPT_STRING_MODELS: frozenset[int] = frozenset({160})

METER_MODELS: frozenset[int] = frozenset({201, 202, 203, 204, 211, 212, 213, 214})
METER_SINGLE_PHASE: frozenset[int] = frozenset({201, 211})
METER_SPLIT_PHASE: frozenset[int] = frozenset({202, 212})
METER_THREE_PHASE_WYE: frozenset[int] = frozenset({203, 213})
METER_THREE_PHASE_DELTA: frozenset[int] = frozenset({204, 214})

DER_700_MODELS: frozenset[int] = frozenset({
    701, 702, 703, 704, 705, 706, 707, 708, 709, 710, 711, 712, 713, 714
})
DER_700_CORE_AC: frozenset[int] = frozenset({701})
DER_700_CAPACITY: frozenset[int] = frozenset({702})
DER_700_CONTROLS: frozenset[int] = frozenset({704, 705, 706})
DER_700_DC_PORTS: frozenset[int] = frozenset({714})

STORAGE_BESS_MODELS: frozenset[int] = frozenset({124, 802, 803, 804, 805})


class SunSpecCliMode(str, Enum):
    """Execution modes for SunSpec definition generation."""
    AUTO = "auto"    # Generate when technically relevant and confirmed
    FORCE = "force"  # Force SunSpec generation using standard models
    OFF = "off"      # Strictly skip SunSpec generation


class SunSpecProfileGeneration(str, Enum):
    """Mutually exclusive SunSpec architectural profiles."""
    LEGACY = "Legacy"  # Models 1 + 101/102/103 + 120-126 + 160
    DER700 = "DER700"  # Models 1 + 701, 702, 704, 714


class SunSpecTopology(str, Enum):
    """AC and DC electrical topologies."""
    SINGLE_PHASE = "SinglePhase"
    SPLIT_PHASE = "SplitPhase"
    THREE_PHASE = "ThreePhase"
    THREE_PHASE_DELTA = "ThreePhaseDelta"
    STORAGE = "Storage"
    CONTROL = "Control"
    UNKNOWN = "General"
