"""
SunSpec Alliance Protocol Module for WebdynSunPM.

Provides comprehensive, typed support for:
  - Official SunSpec Information Models Registry (Legacy & DER 700 series)
  - PICS Document Parsing (Excel PICS, SMDX, JSON)
  - Conditional Relevance Evaluation & Architectural Slicing (SunSpecRelevanceEvaluator)
  - Compliant 11-Column Webdyn Definition Generation (SunSpecDefinitionGenerator)
"""

from DefFileGenerator.sunspec.constants import (
    BASE_ADDRESS_0,
    BASE_ADDRESS_1,
    COMMON_MODEL_LEN_ALIGNED,
    COMMON_MODEL_LEN_STANDARD,
    DER_700_MODELS,
    INVERTER_CONTROL_MODELS,
    LEGACY_INVERTER_MODELS,
    METER_MODELS,
    MPPT_STRING_MODELS,
    SENTINEL_VALUES,
    STORAGE_BESS_MODELS,
    SUNS_END_MODEL_ID,
    SUNS_MAGIC_HEX,
    SUNS_MAGIC_STR,
    SunSpecCliMode,
    SunSpecProfileGeneration,
    SunSpecTopology,
)
from DefFileGenerator.sunspec.evaluator import (
    SunSpecGenerationTarget,
    SunSpecRelevanceDecision,
    SunSpecRelevanceEvaluator,
)
from DefFileGenerator.sunspec.generator import (
    SunSpecDefinitionGenerator,
    SunSpecGenerationResult,
)
from DefFileGenerator.sunspec.models_registry import (
    SUNSPEC_MODELS,
    SUNSPEC_MODELS_REGISTRY,
    SunSpecModelDef,
    SunSpecPoint,
    get_sunspec_model,
    list_supported_models,
)
from DefFileGenerator.sunspec.pics_parser import (
    PICSDocument,
    PICSModel,
    PICSParser,
    PICSPoint,
)

__all__ = [
    "BASE_ADDRESS_0",
    "BASE_ADDRESS_1",
    "COMMON_MODEL_LEN_ALIGNED",
    "COMMON_MODEL_LEN_STANDARD",
    "DER_700_MODELS",
    "INVERTER_CONTROL_MODELS",
    "LEGACY_INVERTER_MODELS",
    "METER_MODELS",
    "MPPT_STRING_MODELS",
    "SENTINEL_VALUES",
    "STORAGE_BESS_MODELS",
    "SUNS_END_MODEL_ID",
    "SUNS_MAGIC_HEX",
    "SUNS_MAGIC_STR",
    "SunSpecCliMode",
    "SunSpecProfileGeneration",
    "SunSpecTopology",
    "SunSpecGenerationTarget",
    "SunSpecRelevanceDecision",
    "SunSpecRelevanceEvaluator",
    "SunSpecDefinitionGenerator",
    "SunSpecGenerationResult",
    "SUNSPEC_MODELS",
    "SUNSPEC_MODELS_REGISTRY",
    "SunSpecModelDef",
    "SunSpecPoint",
    "get_sunspec_model",
    "list_supported_models",
    "PICSDocument",
    "PICSModel",
    "PICSParser",
    "PICSPoint",
]
