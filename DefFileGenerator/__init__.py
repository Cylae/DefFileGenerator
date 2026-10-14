"""WebdynSunPM definition file generator and documentation parser."""

from DefFileGenerator.def_gen import (
    CSVHeaderConfig,
    Generator,
    GeneratorConfig,
    RegisterEntry,
    ValidationIssue,
    ValidationReport,
    WebdynDefConfig,
)
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.knowledge_miner import (
    SUNPM_CANONICAL_TAGS,
    SUNPM_LEGACY_TO_CANONICAL,
    enrich_definition_file,
    enrich_definition_rows,
    infer_sunpm_tag,
    infer_webdyn_tag,
    normalize_tag_to_sunpm,
)

__version__ = "0.2.1"

__all__ = [
    "CSVHeaderConfig",
    "Extractor",
    "Generator",
    "GeneratorConfig",
    "RegisterEntry",
    "SUNPM_CANONICAL_TAGS",
    "SUNPM_LEGACY_TO_CANONICAL",
    "ValidationIssue",
    "ValidationReport",
    "WebdynDefConfig",
    "enrich_definition_file",
    "enrich_definition_rows",
    "infer_sunpm_tag",
    "infer_webdyn_tag",
    "normalize_tag_to_sunpm",
]
