"""
Canonical Intermediate Representation and Typed Schemas for Autonomous Modbus Discovery.
Decouples extraction, normalization, and validation from specific gateway CSV formats.
"""

from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class ProtocolTransport(str, Enum):
    MODBUS_RTU = "modbusRTU"
    MODBUS_TCP = "modbusTCP"
    SUNSPEC_RTU = "sunspecRTU"
    SUNSPEC_TCP = "sunspecTCP"


class RegisterSpace(str, Enum):
    HOLDING_REGISTER = "holding_register"  # FC03 / FC06 / FC16
    INPUT_REGISTER = "input_register"  # FC04
    COIL = "coil"  # FC01 / FC05 / FC15
    DISCRETE_INPUT = "discrete_input"  # FC02


class Endianness(str, Enum):
    BIG_ENDIAN = "BE"  # Standard Modbus (AB / ABCD)
    LITTLE_ENDIAN = "LE"  # Word reversed (BA / CDAB)
    MID_BIG_ENDIAN = "MBE"  # Byte swapped within word
    MID_LITTLE_ENDIAN = "MLE"  # Word swapped, byte preserved


class RegisterDataType(str, Enum):
    U16 = "U16"
    S16 = "S16"
    U32 = "U32"
    S32 = "S32"
    U64 = "U64"
    S64 = "S64"
    FLOAT32 = "FLOAT32"
    FLOAT64 = "FLOAT64"
    STRING = "STRING"
    BITFIELD16 = "BITFIELD16"
    BITFIELD32 = "BITFIELD32"


class ProvenanceRecord(BaseModel):
    document_sha256: str
    document_title: str
    page_number: int | None = None
    table_index: int | None = None
    sheet_name: str | None = None
    cell_range: str | None = None
    raw_text_snippet: str = ""
    extraction_confidence: float = Field(default=1.0, ge=0.0, le=1.0)


class CanonicalRegisterRecord(BaseModel):
    id: str  # Unique internal identifier
    name: str  # Manufacturer variable name
    address: int = Field(ge=0, le=65535)  # Protocol address (0-65535)
    address_base_provenance: str = "documented_0_based"  # "documented_0_based", "inferred_1_based"
    register_space: RegisterSpace = RegisterSpace.HOLDING_REGISTER
    read_fc: int = 3  # 3: Holding, 4: Input, 1: Coil, 2: Discrete
    write_fc: int | None = None  # 6, 16, 5, 15
    data_type: RegisterDataType = RegisterDataType.U16
    register_count: int = Field(default=1, ge=1, le=4)
    endianness: Endianness = Endianness.BIG_ENDIAN
    factor: float = 1.0
    offset: float = 0.0
    unit: str = ""
    canonical_sunpm_tag: str | None = None  # Left blank if uncertified by official spec
    action_code: int = 4  # 1: Alarm, 2: Event, 3: Param, 4: Measurement
    is_writable: bool = False
    is_optional: bool = False
    sentinels: list[int] = Field(default_factory=list)
    bit_allocations: dict[str, int] | None = None
    provenance: ProvenanceRecord | None = None


class CanonicalEquipmentModel(BaseModel):
    manufacturer: str
    model: str
    hardware_revisions: list[str] = Field(default_factory=list)
    supported_firmwares: list[str] = Field(default_factory=list)
    transport: ProtocolTransport = ProtocolTransport.MODBUS_RTU
    category: str = "Inverter"  # Inverter, Meter, Battery, Sensor, Logger, Other
    registers: list[CanonicalRegisterRecord] = Field(default_factory=list)
    sunspec_base_address: int | None = None
    total_registers: int = 0
    unresolved_rows: list[dict[str, Any]] = Field(default_factory=list)


class IdentityCandidate(BaseModel):
    manufacturer: str
    model_family: str
    exact_model: str
    hardware_revision: str | None = None
    firmware_version: str | None = None
    confidence_score: float = Field(default=1.0, ge=0.0, le=1.0)
    is_ambiguous: bool = False
    distinguishing_factors: list[str] = Field(default_factory=list)


class DocumentCandidate(BaseModel):
    source_url: str
    title: str
    publisher: str
    is_official_manufacturer: bool = True
    relevance_rank: int = 1  # 1 to 6 scale
    firmware_applicability: list[str] = Field(default_factory=list)
    document_date: str | None = None
    sha256_preview: str | None = None


class FetchOutcome(str, Enum):
    SUCCESS = "success"
    BLOCKED_SSRF = "blocked_ssrf"
    RATE_LIMITED = "rate_limited"
    OVERSPEC_BYTES = "overspec_bytes"
    ACCESS_RESTRICTED = "access_restricted"
    NETWORK_TIMEOUT = "network_timeout"
    CORRUPT_DOCUMENT = "corrupt_document"


class FetchResult(BaseModel):
    outcome: FetchOutcome
    file_bytes: bytes | None = None
    sha256: str = ""
    content_type: str = ""
    final_url: str = ""
    error_message: str | None = None


class ValidationStatus(str, Enum):
    VALIDATED_OFFLINE = "VALIDATED_OFFLINE"
    REVIEW_REQUIRED = "REVIEW_REQUIRED"
    AMBIGUOUS_MODEL = "AMBIGUOUS_MODEL"
    SOURCE_CONFLICT = "SOURCE_CONFLICT"
    NOT_FOUND = "NOT_FOUND"
    BLOCKED_ACCESS = "BLOCKED_ACCESS"
    FAILED_VALIDATION = "FAILED_VALIDATION"


class ValidationOutcome(BaseModel):
    status: ValidationStatus
    pillar_statuses: dict[str, bool] = Field(default_factory=dict)
    blocking_issues: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    reconciliation: dict[str, int] = Field(default_factory=dict)


class ExportTarget(str, Enum):
    WEBDYN_SUNPM = "webdyn_sunpm"
    WEBDYN_CLASSIC = "webdyn_classic"
