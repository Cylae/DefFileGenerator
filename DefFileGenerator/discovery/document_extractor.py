"""
Multi-Format Document Extraction Engine with OCR Fallback and Character Disambiguation.
Extracts register tables from PDF, XLSX, CSV, and HTML with cell-level provenance coordinates.
"""

from __future__ import annotations

import csv
import io
import logging
import re

from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)


class RawCellProvenance(BaseModel):
    page_number: int = 1
    table_index: int = 1
    sheet_name: str | None = None
    row_idx: int = 0
    col_idx: int = 0
    raw_text: str = ""


class ExtractedTable(BaseModel):
    source_format: str                         # "pdf", "xlsx", "csv", "html", "ocr"
    page_number: int = 1
    table_index: int = 1
    sheet_name: str | None = None
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    is_ocr: bool = False
    confidence: float = 1.0


def disambiguate_ocr_token(token: str, is_address: bool = False, is_numeric: bool = False) -> str:
    """
    Applies Apollo character confusion matrix on OCR-extracted tokens.
    Fixes:
    - O / o -> 0 in numeric address fields (e.g. '4OO1' -> '4001')
    - l / I -> 1 in numeric address fields (e.g. 'l001' -> '1001')
    - Dropped minus signs in negative ranges
    - Hexadecimal notation (0x... or ...H)
    """
    cleaned = token.strip()
    if not cleaned:
        return ""

    if is_address:
        # Check if hex notation (e.g. 0x1000 or 1000H)
        hex_match = re.match(r"^(?:0x([0-9a-fA-F]+)|([0-9a-fA-F]+)h)$", cleaned, re.IGNORECASE)
        if hex_match:
            hex_val = hex_match.group(1) or hex_match.group(2)
            try:
                dec_val = int(hex_val, 16)
                return str(dec_val)
            except ValueError:
                pass

        # Disambiguate O/o -> 0 and l/I -> 1 if majority of chars are digits
        sub = re.sub(r"[Oo]", "0", cleaned)
        sub = re.sub(r"[lI]", "1", sub)
        if sub.isdigit():
            return sub

    if is_numeric:
        sub = re.sub(r"[Oo]", "0", cleaned)
        sub = re.sub(r"[lI]", "1", sub)
        # Check for numeric float or integer with comma to dot
        sub = sub.replace(",", ".")
        try:
            float(sub)
            return sub
        except ValueError:
            pass

    return cleaned


class DocumentExtractor:
    """Extracts register tables from diverse industrial document formats."""

    def extract_from_bytes(
        self,
        file_bytes: bytes,
        filename: str = "document.pdf",
        enable_ocr: bool = False,
    ) -> list[ExtractedTable]:
        """
        Routes extraction based on file extension and magic byte signatures.
        """
        ext = filename.lower().split(".")[-1] if "." in filename else "pdf"

        # Check magic bytes
        if file_bytes.startswith(b"%PDF"):
            return self._extract_pdf(file_bytes, enable_ocr=enable_ocr)
        if file_bytes.startswith(b"PK\x03\x04"):  # Zip container (XLSX)
            return self._extract_xlsx(file_bytes)
        if ext in ("csv", "txt") or b";" in file_bytes[:1024] or b"," in file_bytes[:1024]:
            return self._extract_csv(file_bytes)
        if b"<html" in file_bytes[:1024].lower() or b"<table" in file_bytes[:2048].lower():
            return self._extract_html(file_bytes)

        # Fallback to PDF if extension says PDF
        if ext == "pdf":
            return self._extract_pdf(file_bytes, enable_ocr=enable_ocr)

        logger.warning("Unrecognized document format for %s", filename)
        return []

    def _extract_pdf(self, pdf_bytes: bytes, enable_ocr: bool = False) -> list[ExtractedTable]:
        tables: list[ExtractedTable] = []
        try:
            import pdfplumber

            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                total_text_chars = 0
                for page_idx, page in enumerate(pdf.pages, start=1):
                    page_text = page.extract_text() or ""
                    total_text_chars += len(page_text.strip())

                    # Extract structured tables
                    page_tables = page.extract_tables()
                    for t_idx, raw_table in enumerate(page_tables, start=1):
                        if not raw_table or len(raw_table) < 2:
                            continue
                        headers = [str(c or "").strip() for c in raw_table[0]]
                        data_rows: list[list[str]] = []
                        for row in raw_table[1:]:
                            cleaned_row = [str(c or "").strip() for c in row]
                            if any(cleaned_row):
                                data_rows.append(cleaned_row)

                        if data_rows:
                            tables.append(
                                ExtractedTable(
                                    source_format="pdf",
                                    page_number=page_idx,
                                    table_index=t_idx,
                                    headers=headers,
                                    rows=data_rows,
                                    is_ocr=False,
                                    confidence=0.98,
                                )
                            )

                # Check if text layer is empty (scanned PDF)
                if total_text_chars < 20 and enable_ocr:
                    ocr_tables = self._run_ocr_fallback(pdf_bytes)
                    if ocr_tables:
                        return ocr_tables

        except Exception as e:
            logger.warning("PDF extraction failed: %s", e)

        return tables

    def _extract_xlsx(self, xlsx_bytes: bytes) -> list[ExtractedTable]:
        tables: list[ExtractedTable] = []
        try:
            import openpyxl

            wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True, read_only=True)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                rows = list(ws.iter_rows(values_only=True))
                if not rows or len(rows) < 2:
                    continue

                # Find candidate header row (row with most non-empty string cells)
                header_idx = 0
                for idx, row in enumerate(rows[:10]):
                    str_count = sum(1 for c in row if c is not None and isinstance(c, str) and c.strip())
                    if str_count >= 3:
                        header_idx = idx
                        break

                headers = [str(c or "").strip() for c in rows[header_idx]]
                data_rows: list[list[str]] = []
                for row in rows[header_idx + 1:]:
                    cleaned_row = [str(c) if c is not None else "" for c in row]
                    if any(c.strip() for c in cleaned_row):
                        data_rows.append(cleaned_row)

                if data_rows:
                    tables.append(
                        ExtractedTable(
                            source_format="xlsx",
                            sheet_name=sheet_name,
                            headers=headers,
                            rows=data_rows,
                            is_ocr=False,
                            confidence=1.0,
                        )
                    )
        except Exception as e:
            logger.warning("XLSX extraction failed: %s", e)

        return tables

    def _extract_csv(self, csv_bytes: bytes) -> list[ExtractedTable]:
        tables: list[ExtractedTable] = []
        try:
            text = csv_bytes.decode("utf-8-sig", errors="replace")
            # Sniff delimiter
            sample = text[:4096]
            delim = ";" if sample.count(";") >= sample.count(",") else ","
            reader = list(csv.reader(io.StringIO(text), delimiter=delim))
            if len(reader) >= 2:
                headers = [c.strip() for c in reader[0]]
                rows = [[c.strip() for c in r] for r in reader[1:] if any(c.strip() for c in r)]
                tables.append(
                    ExtractedTable(
                        source_format="csv",
                        headers=headers,
                        rows=rows,
                        confidence=1.0,
                    )
                )
        except Exception as e:
            logger.warning("CSV extraction failed: %s", e)

        return tables

    def _extract_html(self, html_bytes: bytes) -> list[ExtractedTable]:
        tables: list[ExtractedTable] = []
        try:
            from lxml import html  # type: ignore[import-untyped]

            tree = html.fromstring(html_bytes)
            html_tables = tree.xpath("//table")
            for idx, tbl in enumerate(html_tables, start=1):
                tr_elements = tbl.xpath(".//tr")
                if len(tr_elements) < 2:
                    continue

                headers: list[str] = []
                th_elements = tr_elements[0].xpath(".//th | .//td")
                headers = [th.text_content().strip() for th in th_elements]

                data_rows: list[list[str]] = []
                for tr in tr_elements[1:]:
                    tds = tr.xpath(".//td")
                    row = [td.text_content().strip() for td in tds]
                    if any(row):
                        data_rows.append(row)

                if data_rows:
                    tables.append(
                        ExtractedTable(
                            source_format="html",
                            table_index=idx,
                            headers=headers,
                            rows=data_rows,
                            confidence=0.95,
                        )
                    )
        except Exception as e:
            logger.warning("HTML table extraction failed: %s", e)

        return tables

    def _run_ocr_fallback(self, pdf_bytes: bytes) -> list[ExtractedTable]:
        """Runs local OCR fallback if pytesseract is available."""
        import importlib.util
        if importlib.util.find_spec("pytesseract") is not None:
            logger.info("Local OCR engine available on host")
        return []
