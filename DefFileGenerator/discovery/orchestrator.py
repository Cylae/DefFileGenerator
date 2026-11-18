"""
Workflow Orchestrator for Autonomous Modbus Discovery.
Coordinates end-to-end execution across identity resolution, source discovery,
secure document fetching, extraction, normalization, validation, and export.
"""

from __future__ import annotations

import logging

from pydantic import BaseModel

from DefFileGenerator.discovery.definition_exporter import DefinitionExporter
from DefFileGenerator.discovery.document_extractor import DocumentExtractor
from DefFileGenerator.discovery.document_fetcher import DocumentFetcher
from DefFileGenerator.discovery.evidence_reporter import EvidenceReporter
from DefFileGenerator.discovery.identity_resolver import IdentityResolver
from DefFileGenerator.discovery.models import (
    CanonicalEquipmentModel,
    DocumentCandidate,
    ExportTarget,
    FetchOutcome,
    FetchResult,
    IdentityCandidate,
    ValidationOutcome,
    ValidationStatus,
)
from DefFileGenerator.discovery.register_normalizer import RegisterNormalizer
from DefFileGenerator.discovery.semantic_validator import SemanticValidator
from DefFileGenerator.discovery.source_discovery import SourceDiscovery

logger = logging.getLogger(__name__)


class AutonomousPipelineResult(BaseModel):
    success: bool
    status: ValidationStatus
    query: str
    identity: IdentityCandidate | None = None
    document: DocumentCandidate | None = None
    fetch: FetchResult | None = None
    model: CanonicalEquipmentModel | None = None
    validation: ValidationOutcome | None = None
    csv_content: str | None = None
    evidence_report: str | None = None
    error_message: str | None = None


class WorkflowOrchestrator:
    """Coordinates autonomous Modbus definition discovery and generation."""

    def __init__(
        self,
        resolver: IdentityResolver | None = None,
        discovery: SourceDiscovery | None = None,
        fetcher: DocumentFetcher | None = None,
        extractor: DocumentExtractor | None = None,
        normalizer: RegisterNormalizer | None = None,
        validator: SemanticValidator | None = None,
        exporter: DefinitionExporter | None = None,
        reporter: EvidenceReporter | None = None,
    ) -> None:
        self.resolver = resolver or IdentityResolver()
        self.discovery = discovery or SourceDiscovery()
        self.fetcher = fetcher or DocumentFetcher()
        self.extractor = extractor or DocumentExtractor()
        self.normalizer = normalizer or RegisterNormalizer()
        self.validator = validator or SemanticValidator()
        self.exporter = exporter or DefinitionExporter()
        self.reporter = reporter or EvidenceReporter()

    async def run(
        self,
        query: str,
        target: ExportTarget = ExportTarget.WEBDYN_SUNPM,
        enable_ocr: bool = True,
        override_file_bytes: bytes | None = None,
        override_filename: str = "manual.pdf",
    ) -> AutonomousPipelineResult:
        """
        Executes autonomous pipeline from reference query to target CSV.
        Supports manual document override fallback when automated retrieval is insufficient.
        """
        clean_query = query.strip()
        if not clean_query:
            return AutonomousPipelineResult(
                success=False,
                status=ValidationStatus.NOT_FOUND,
                query=query,
                error_message="Empty equipment reference query provided.",
            )

        # 1. Identity Resolution
        candidates = self.resolver.resolve(clean_query)
        if not candidates:
            return AutonomousPipelineResult(
                success=False,
                status=ValidationStatus.NOT_FOUND,
                query=query,
                error_message=f"Could not identify manufacturer or model for query '{query}'.",
            )

        identity = candidates[0]
        if identity.is_ambiguous:
            return AutonomousPipelineResult(
                success=False,
                status=ValidationStatus.AMBIGUOUS_MODEL,
                query=query,
                identity=identity,
                error_message=f"Ambiguous equipment model. Possible variants: {', '.join(identity.distinguishing_factors)}",
            )

        # 2. Source Document Acquisition (Override or Network Discovery)
        doc_candidate: DocumentCandidate
        fetch_result: FetchResult

        if override_file_bytes:
            import hashlib

            sha256_hash = hashlib.sha256(override_file_bytes).hexdigest()
            doc_candidate = DocumentCandidate(
                source_url=f"local://manual_upload/{override_filename}",
                title=f"Manual Upload for {identity.manufacturer} {identity.exact_model}",
                publisher="User Provided",
                relevance_rank=1,
            )
            fetch_result = FetchResult(
                outcome=FetchOutcome.SUCCESS,
                file_bytes=override_file_bytes,
                sha256=sha256_hash,
                content_type="application/pdf",
                final_url=doc_candidate.source_url,
            )
        else:
            sources = await self.discovery.discover(identity)
            if not sources:
                return AutonomousPipelineResult(
                    success=False,
                    status=ValidationStatus.NOT_FOUND,
                    query=query,
                    identity=identity,
                    error_message=f"No authoritative documentation source found for {identity.manufacturer} {identity.exact_model}.",
                )
            doc_candidate = sources[0]
            fetch_result = await self.fetcher.fetch(doc_candidate)

            if fetch_result.outcome != FetchOutcome.SUCCESS or not fetch_result.file_bytes:
                status = (
                    ValidationStatus.BLOCKED_ACCESS
                    if fetch_result.outcome
                    in (FetchOutcome.BLOCKED_SSRF, FetchOutcome.ACCESS_RESTRICTED)
                    else ValidationStatus.NOT_FOUND
                )
                return AutonomousPipelineResult(
                    success=False,
                    status=status,
                    query=query,
                    identity=identity,
                    document=doc_candidate,
                    fetch=fetch_result,
                    error_message=f"Failed to retrieve document: {fetch_result.error_message}",
                )

        # 3. Document Extraction
        assert fetch_result.file_bytes is not None
        extracted_tables = self.extractor.extract_from_bytes(
            fetch_result.file_bytes,
            filename=override_filename,
            enable_ocr=enable_ocr,
        )
        if not extracted_tables:
            return AutonomousPipelineResult(
                success=False,
                status=ValidationStatus.FAILED_VALIDATION,
                query=query,
                identity=identity,
                document=doc_candidate,
                fetch=fetch_result,
                error_message="Document contained no parseable register tables.",
            )

        # 4. Modbus Register Normalization
        canonical_model = self.normalizer.normalize(
            tables=extracted_tables,
            identity=identity,
            document_sha256=fetch_result.sha256,
            document_title=doc_candidate.title,
        )

        # 5. Semantic Validation (5 Apollo Pillars)
        validation = self.validator.validate(canonical_model)
        if validation.status == ValidationStatus.FAILED_VALIDATION:
            return AutonomousPipelineResult(
                success=False,
                status=validation.status,
                query=query,
                identity=identity,
                document=doc_candidate,
                fetch=fetch_result,
                model=canonical_model,
                validation=validation,
                error_message=f"Validation failed: {', '.join(validation.blocking_issues)}",
            )

        # 6. Target Export
        csv_content = self.exporter.export(canonical_model, target=target)

        # 7. Evidence Report Generation
        evidence_report = self.reporter.generate_markdown_report(
            model=canonical_model,
            document_candidate=doc_candidate,
            fetch_result=fetch_result,
            validation=validation,
        )

        return AutonomousPipelineResult(
            success=True,
            status=validation.status,
            query=query,
            identity=identity,
            document=doc_candidate,
            fetch=fetch_result,
            model=canonical_model,
            validation=validation,
            csv_content=csv_content,
            evidence_report=evidence_report,
        )
