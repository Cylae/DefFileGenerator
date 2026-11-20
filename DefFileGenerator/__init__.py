"""WebdynSunPM definition file generator and documentation parser."""

from DefFileGenerator._version import __version__
from DefFileGenerator.batch_parser import (
    BatchModbusParser,
    batch_parse,
)
from DefFileGenerator.def_gen import (
    CSVHeaderConfig,
    Generator,
    GeneratorConfig,
    RegisterEntry,
    ValidationIssue,
    ValidationReport,
    WebdynDefConfig,
)
from DefFileGenerator.dyness_generator import (
    DynessDefinitionGenerator,
    DynessProtocolParser,
    generate_dyness_definitions,
)
from DefFileGenerator.excel_modbus_extractor import (
    ExcelModbusExtractor,
    extract_excel_modbus,
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
from DefFileGenerator.pdf_parser import (
    PDFModbusParser,
    parse_pdf_modbus,
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

__all__ = [
    "__version__",
    "BatchModbusParser",
    "CSVHeaderConfig",
    "DatasetKnowledge",
    "DefinitionAuditIssue",
    "DefinitionAuditReport",
    "DefinitionFileStats",
    "DirectoryReportStats",
    "DynessDefinitionGenerator",
    "DynessProtocolParser",
    "ExcelModbusExtractor",
    "Extractor",
    "Generator",
    "GeneratorConfig",
    "PDFModbusParser",
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
    "batch_parse",
    "enrich_definition_file",
    "enrich_definition_rows",
    "export_report_json",
    "extract_excel_modbus",
    "generate_directory_technical_report",
    "generate_dyness_definitions",
    "generate_single_definition_report",
    "infer_sunpm_tag",
    "infer_webdyn_tag",
    "is_canonical_sunpm_tag",
    "normalize_tag_to_sunpm",
    "parse_pdf_modbus",
    "train_knowledge_from_definitions",
]
