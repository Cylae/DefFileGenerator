"""
Autonomous Modbus Discovery & Definition Generation Package.
Implements identity resolution, controlled secure document retrieval,
canonical intermediate representation, Modbus semantic normalization,
and target-certified definition compilation.
"""

from DefFileGenerator.discovery.models import (
    CanonicalEquipmentModel,
    CanonicalRegisterRecord,
    DocumentCandidate,
    Endianness,
    FetchOutcome,
    FetchResult,
    IdentityCandidate,
    ProtocolTransport,
    ProvenanceRecord,
    RegisterDataType,
    RegisterSpace,
    ValidationOutcome,
    ValidationStatus,
)

__all__ = [
    "CanonicalEquipmentModel",
    "CanonicalRegisterRecord",
    "DocumentCandidate",
    "Endianness",
    "FetchOutcome",
    "FetchResult",
    "IdentityCandidate",
    "ProtocolTransport",
    "ProvenanceRecord",
    "RegisterDataType",
    "RegisterSpace",
    "ValidationOutcome",
    "ValidationStatus",
]
