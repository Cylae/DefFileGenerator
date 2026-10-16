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
    DatasetKnowledge,
    DefinitionAuditIssue,
    DefinitionAuditReport,
    audit_definition_file,
    audit_definitions_directory,
    auto_fix_definition_file,
    enrich_definition_file,
    enrich_definition_rows,
    infer_sunpm_tag,
    infer_webdyn_tag,
    is_canonical_sunpm_tag,
    normalize_tag_to_sunpm,
    train_knowledge_from_definitions,
)
from DefFileGenerator.report_generator import (
    DefinitionFileStats,
    DirectoryReportStats,
    analyze_definition_file,
    analyze_definitions_directory,
    export_report_json,
    generate_directory_technical_report,
    generate_single_definition_report,
)

__version__ = "0.2.1"

__all__ = [
    "CSVHeaderConfig",
    "DatasetKnowledge",
    "DefinitionAuditIssue",
    "DefinitionAuditReport",
    "DefinitionFileStats",
    "DirectoryReportStats",
    "Extractor",
    "Generator",
    "GeneratorConfig",
    "RegisterEntry",
    "SUNPM_CANONICAL_TAGS",
    "SUNPM_LEGACY_TO_CANONICAL",
    "ValidationIssue",
    "ValidationReport",
    "WebdynDefConfig",
    "analyze_definition_file",
    "analyze_definitions_directory",
    "audit_definition_file",
    "audit_definitions_directory",
    "auto_fix_definition_file",
    "enrich_definition_file",
    "enrich_definition_rows",
    "export_report_json",
    "generate_directory_technical_report",
    "generate_single_definition_report",
    "infer_sunpm_tag",
    "infer_webdyn_tag",
    "is_canonical_sunpm_tag",
    "normalize_tag_to_sunpm",
    "train_knowledge_from_definitions",
]

