"""
Batch Modbus Parser Module.

Orchestrates batch extraction and definition generation across directories containing heterogeneous
vendor documentation files (.xlsx, .xls, .pdf, .csv, .xml, .json).
Features multi-worker concurrency, vendor/model metadata inference, error containment, and Markdown reporting.
"""

from __future__ import annotations

import concurrent.futures
import logging
import time
from pathlib import Path
from typing import Any

from DefFileGenerator.main import ALLOWED_EXTENSIONS, batch_process_file
from DefFileGenerator.report_generator import generate_directory_technical_report


class BatchModbusParser:
    """
    High-level orchestrator for batch processing directories of Modbus documentation.
    """

    def __init__(
        self,
        protocol: str = "modbusRTU",
        category: str | None = None,
        manufacturer: str | None = None,
        tag: bool = True,
        strict: bool = False,
        force: bool = False,
        workers: int = 1,
    ) -> None:
        """
        Initializes the BatchModbusParser.

        Args:
            protocol: Target Webdyn protocol identifier (default: "modbusRTU").
            category: Optional default device category.
            manufacturer: Optional default manufacturer override.
            tag: Whether to infer and enrich SunPM telemetry tags.
            strict: Whether to enforce strict 5 pillars validation.
            force: Whether to overwrite existing destination files.
            workers: Concurrency worker thread count (default: 1).
        """
        self.protocol = protocol
        self.category = category
        self.manufacturer = manufacturer
        self.tag = tag
        self.strict = strict
        self.force = force
        self.workers = max(1, workers)

    def discover_files(self, input_dir: str | Path) -> list[Path]:
        """
        Scans input_dir recursively or flat for supported documentation files.

        Args:
            input_dir: Directory to scan.

        Returns:
            list[Path]: Discovered file paths matching ALLOWED_EXTENSIONS.
        """
        p = Path(input_dir)
        if not p.is_dir():
            return []
        files = [
            f
            for f in sorted(p.iterdir())
            if f.is_file() and f.suffix.lower() in ALLOWED_EXTENSIONS and not f.name.startswith("~")
        ]
        return files

    def process_directory(
        self,
        input_dir: str | Path,
        output_dir: str | Path,
        report_path: str | Path | None = None,
    ) -> dict[str, Any]:
        """
        Processes all candidate documentation files in input_dir and writes definitions to output_dir.

        Args:
            input_dir: Directory containing raw vendor manuals and spreadsheets.
            output_dir: Directory to store generated Webdyn definition CSVs.
            report_path: Optional file path to output a technical Markdown report.

        Returns:
            dict[str, Any]: Execution summary dictionary with counts, errors, and timing.
        """
        start_time = time.monotonic()
        in_path = Path(input_dir)
        out_path = Path(output_dir)

        if not in_path.is_dir():
            raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

        out_path.mkdir(parents=True, exist_ok=True)
        files = self.discover_files(in_path)

        summary: dict[str, Any] = {
            "total_files": len(files),
            "processed": 0,
            "successes": 0,
            "failures": 0,
            "skipped": 0,
            "duration_seconds": 0.0,
            "files": [],
            "errors": [],
        }

        if not files:
            logging.warning(f"No supported documentation files found in {input_dir}")
            summary["duration_seconds"] = round(time.monotonic() - start_time, 3)
            return summary

        tasks: list[dict[str, Any]] = []
        for file_path in files:
            task = {
                "file_path": file_path,
                "output_dir": out_path,
                "protocol": self.protocol,
                "category": self.category,
                "manufacturer": self.manufacturer,
                "tag": self.tag,
                "strict": self.strict,
                "force": self.force,
            }
            tasks.append(task)

        if self.workers > 1:
            with concurrent.futures.ThreadPoolExecutor(max_workers=self.workers) as executor:
                future_to_task = {
                    executor.submit(
                        batch_process_file,
                        t["file_path"],
                        t["output_dir"],
                        t["protocol"],
                        t["category"],
                        t["manufacturer"],
                        t["tag"],
                        t["strict"],
                        t["force"],
                    ): t
                    for t in tasks
                }
                for future in concurrent.futures.as_completed(future_to_task):
                    t = future_to_task[future]
                    fname = t["file_path"].name
                    try:
                        res = future.result()
                        summary["processed"] += 1
                        if res.get("status") == "success":
                            summary["successes"] += 1
                            summary["files"].append(
                                {"name": fname, "status": "success", "out": str(res.get("output"))}
                            )
                        elif res.get("status") == "skipped":
                            summary["skipped"] += 1
                            summary["files"].append({"name": fname, "status": "skipped"})
                        else:
                            summary["failures"] += 1
                            err_msg = res.get("error", "Unknown error")
                            summary["errors"].append({"name": fname, "error": err_msg})
                            summary["files"].append(
                                {"name": fname, "status": "failed", "error": err_msg}
                            )
                    except Exception as exc:
                        summary["processed"] += 1
                        summary["failures"] += 1
                        summary["errors"].append({"name": fname, "error": str(exc)})
                        summary["files"].append(
                            {"name": fname, "status": "failed", "error": str(exc)}
                        )
        else:
            for t in tasks:
                fname = t["file_path"].name
                try:
                    res = batch_process_file(
                        t["file_path"],
                        t["output_dir"],
                        t["protocol"],
                        t["category"],
                        t["manufacturer"],
                        t["tag"],
                        t["strict"],
                        t["force"],
                    )
                    summary["processed"] += 1
                    if res.get("status") == "success":
                        summary["successes"] += 1
                        summary["files"].append(
                            {"name": fname, "status": "success", "out": str(res.get("output"))}
                        )
                    elif res.get("status") == "skipped":
                        summary["skipped"] += 1
                        summary["files"].append({"name": fname, "status": "skipped"})
                    else:
                        summary["failures"] += 1
                        err_msg = res.get("error", "Unknown error")
                        summary["errors"].append({"name": fname, "error": err_msg})
                        summary["files"].append(
                            {"name": fname, "status": "failed", "error": err_msg}
                        )
                except Exception as exc:
                    summary["processed"] += 1
                    summary["failures"] += 1
                    summary["errors"].append({"name": fname, "error": str(exc)})
                    summary["files"].append({"name": fname, "status": "failed", "error": str(exc)})

        summary["duration_seconds"] = round(time.monotonic() - start_time, 3)

        if report_path:
            try:
                rep_content = generate_directory_technical_report(out_path)
                Path(report_path).write_text(rep_content, encoding="utf-8")
                summary["report_path"] = str(report_path)
            except Exception as exc:
                logging.warning(f"Failed to generate technical report at {report_path}: {exc}")

        return summary


def batch_parse(
    input_dir: str | Path,
    output_dir: str | Path,
    report_path: str | Path | None = None,
    protocol: str = "modbusRTU",
    category: str | None = None,
    manufacturer: str | None = None,
    tag: bool = True,
    strict: bool = False,
    force: bool = True,
    workers: int = 1,
) -> dict[str, Any]:
    """
    Convenience function to run a batch parsing run across a documentation directory.

    Args:
        input_dir: Directory containing input docs.
        output_dir: Directory to output definitions.
        report_path: Optional markdown report file path.
        protocol: Webdyn protocol string.
        category: Equipment category.
        manufacturer: Equipment manufacturer.
        tag: Enable semantic tagging.
        strict: Enforce strict certification checks.
        force: Overwrite existing outputs.
        workers: Parallel worker count.

    Returns:
        dict[str, Any]: Batch summary results.
    """
    parser = BatchModbusParser(
        protocol=protocol,
        category=category,
        manufacturer=manufacturer,
        tag=tag,
        strict=strict,
        force=force,
        workers=workers,
    )
    return parser.process_directory(input_dir, output_dir, report_path=report_path)
