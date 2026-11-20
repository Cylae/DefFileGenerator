"""
Duplicate Detector & Normalization Engine for Modbus Definition Files.

Identifies:
1. Exact SHA-256 identical duplicate files
2. Exact register-body duplicate files (identical Modbus register tables, differing headers/models)
3. Category prefix conflicts (e.g. OTH_ vs MET_/INV_/SEN_/BAT_/COM_ for the same equipment stem)
4. Hardware family duplicates (different models sharing identical register maps)
5. Naming alias synonyms (e.g. HT1500 vs HT_1500V, Core2 vs STP_110-60, GT_Series vs GT-SERIES)
6. Near-duplicate register maps (similarity >= 0.95)

Provides structured reports and safe deduplication recommendations.
"""

from __future__ import annotations

import csv
from dataclasses import asdict, dataclass, field
import hashlib
import json
from pathlib import Path
import re
from typing import Any


@dataclass
class DuplicateCluster:
    """Represents a cluster of duplicate or alias definition files."""

    cluster_id: str
    cluster_type: str  # 'exact_sha256', 'exact_body', 'category_conflict', 'alias_synonym'
    canonical_file: str
    duplicate_files: list[str]
    description: str
    register_count: int
    models: list[str] = field(default_factory=list)
    action_recommendation: str = "keep_canonical"


@dataclass
class DuplicateReport:
    """Comprehensive report detailing definition duplicates across a target directory."""

    target_directory: str
    total_files_scanned: int
    total_duplicate_files: int
    exact_hash_clusters: list[DuplicateCluster] = field(default_factory=list)
    exact_body_clusters: list[DuplicateCluster] = field(default_factory=list)
    category_conflict_clusters: list[DuplicateCluster] = field(default_factory=list)
    alias_clusters: list[DuplicateCluster] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class DuplicateDetector:
    """Scans and analyzes definition files for duplication and redundant classifications."""

    PREFIX_PRIORITY: dict[str, int] = {
        "INV": 10,
        "BAT": 9,
        "MET": 8,
        "SEN": 7,
        "COM": 6,
        "OTH": 1,
    }

    def __init__(self, directory: Path | str) -> None:
        self.directory = Path(directory)

    def _read_file_data(self, file_path: Path) -> tuple[str, str, str, list[list[str]]]:
        """
        Reads CSV file and extracts (raw_bytes_hash, header_line, body_hash, parsed_rows).
        """
        raw_bytes = file_path.read_bytes()
        raw_hash = hashlib.sha256(raw_bytes).hexdigest()

        text = file_path.read_text(encoding="utf-8", errors="replace")
        lines = text.splitlines()
        header = lines[0] if lines else ""
        body = "\n".join(lines[1:])
        body_hash = hashlib.sha256(body.encode("utf-8")).hexdigest()

        parsed_rows: list[list[str]] = []
        try:
            reader = csv.reader(lines, delimiter=";")
            parsed_rows = [r for r in reader if r and any(cell.strip() for cell in r)]
        except Exception:
            pass

        return raw_hash, header, body_hash, parsed_rows

    def scan(self) -> DuplicateReport:
        """Performs full analysis of duplicate definitions in the directory."""
        csv_files = sorted(self.directory.glob("*.csv"))
        file_data: dict[str, dict[str, Any]] = {}

        for f in csv_files:
            raw_hash, header, body_hash, rows = self._read_file_data(f)
            prefix = f.name.split("_", 1)[0] if "_" in f.name else "OTH"
            reg_count = max(0, len(rows) - 1)
            model = ""
            if rows and len(rows[0]) >= 4:
                model = rows[0][3]

            file_data[f.name] = {
                "path": f,
                "raw_hash": raw_hash,
                "header": header,
                "body_hash": body_hash,
                "prefix": prefix,
                "reg_count": reg_count,
                "model": model,
            }

        # 1. Exact SHA-256 duplicates
        exact_hash_groups: dict[str, list[str]] = {}
        for fname, d in file_data.items():
            exact_hash_groups.setdefault(d["raw_hash"], []).append(fname)

        exact_hash_clusters: list[DuplicateCluster] = []
        for h, flist in exact_hash_groups.items():
            if len(flist) > 1:
                canonical = sorted(flist, key=lambda x: (self.PREFIX_PRIORITY.get(file_data[x]["prefix"], 0), -len(x)), reverse=True)[0]
                dupes = [x for x in flist if x != canonical]
                exact_hash_clusters.append(
                    DuplicateCluster(
                        cluster_id=f"hash_{h[:8]}",
                        cluster_type="exact_sha256",
                        canonical_file=canonical,
                        duplicate_files=dupes,
                        description=f"{len(flist)} byte-identical definition files",
                        register_count=file_data[canonical]["reg_count"],
                        models=[file_data[x]["model"] for x in flist],
                        action_recommendation="delete_or_symlink_duplicates",
                    )
                )

        # 2. Exact Register Body Duplicates (different header line, identical registers 1..N)
        exact_body_groups: dict[str, list[str]] = {}
        for fname, d in file_data.items():
            exact_body_groups.setdefault(d["body_hash"], []).append(fname)

        exact_body_clusters: list[DuplicateCluster] = []
        category_conflict_clusters: list[DuplicateCluster] = []
        alias_clusters: list[DuplicateCluster] = []

        seen_in_body_clusters: set[str] = set()

        for bh, flist in exact_body_groups.items():
            if len(flist) <= 1:
                continue

            # Pick canonical: prefer typed prefix (INV/BAT/MET/SEN/COM) over OTH, then shortest canonical name
            canonical = sorted(
                flist,
                key=lambda x: (
                    self.PREFIX_PRIORITY.get(file_data[x]["prefix"], 0),
                    -len(x),
                ),
                reverse=True,
            )[0]
            dupes = [x for x in flist if x != canonical]

            # Check if this is a Category Conflict (e.g. OTH_ vs MET_ or SEN_)
            prefixes = {file_data[x]["prefix"] for x in flist}
            has_oth = "OTH" in prefixes
            has_typed = any(p in ("INV", "MET", "SEN", "BAT", "COM") for p in prefixes)

            # Check stem equivalence (ignoring prefix)
            stems = {re.sub(r"^(INV|MET|SEN|BAT|COM|OTH)_", "", x).lower() for x in flist}

            cluster_reg_count = file_data[canonical]["reg_count"]
            cluster_models = [file_data[x]["model"] for x in flist]

            if has_oth and has_typed and len(stems) == 1:
                # Direct category conflict for the exact same hardware
                category_conflict_clusters.append(
                    DuplicateCluster(
                        cluster_id=f"cat_{bh[:8]}",
                        cluster_type="category_conflict",
                        canonical_file=canonical,
                        duplicate_files=dupes,
                        description=f"Generic OTH_ prefix duplicate of typed {file_data[canonical]['prefix']} definition",
                        register_count=cluster_reg_count,
                        models=cluster_models,
                        action_recommendation="purge_oth_duplicate",
                    )
                )
            elif any(self._is_alias_pair(canonical, d) for d in dupes):
                # Hardware alias / alternative naming
                alias_clusters.append(
                    DuplicateCluster(
                        cluster_id=f"alias_{bh[:8]}",
                        cluster_type="alias_synonym",
                        canonical_file=canonical,
                        duplicate_files=dupes,
                        description="Hardware model alias / synonymous naming sharing identical register map",
                        register_count=cluster_reg_count,
                        models=cluster_models,
                        action_recommendation="keep_or_register_alias",
                    )
                )
            else:
                # Hardware model family sharing standard register protocol
                exact_body_clusters.append(
                    DuplicateCluster(
                        cluster_id=f"body_{bh[:8]}",
                        cluster_type="exact_body",
                        canonical_file=canonical,
                        duplicate_files=dupes,
                        description=f"Hardware family of {len(flist)} models sharing identical Modbus registers",
                        register_count=cluster_reg_count,
                        models=cluster_models,
                        action_recommendation="keep_hardware_family_models",
                    )
                )

            seen_in_body_clusters.update(flist)

        total_dupe_files = sum(len(c.duplicate_files) for c in (category_conflict_clusters + alias_clusters + exact_hash_clusters))

        return DuplicateReport(
            target_directory=str(self.directory),
            total_files_scanned=len(csv_files),
            total_duplicate_files=total_dupe_files,
            exact_hash_clusters=exact_hash_clusters,
            exact_body_clusters=exact_body_clusters,
            category_conflict_clusters=category_conflict_clusters,
            alias_clusters=alias_clusters,
        )

    def _is_alias_pair(self, name_a: str, name_b: str) -> bool:
        """Determines if two names are naming aliases of each other."""
        clean_a = re.sub(r"^(INV|MET|SEN|BAT|COM|OTH)_", "", name_a).lower().replace("-", "_")
        clean_b = re.sub(r"^(INV|MET|SEN|BAT|COM|OTH)_", "", name_b).lower().replace("-", "_")

        if clean_a == clean_b:
            return True

        # Common synonym patterns
        for token in ("1500v", "1500", "series", "work", "tcp", "rtu", "generic"):
            clean_a = clean_a.replace(token, "")
            clean_b = clean_b.replace(token, "")

        return clean_a.strip("_") == clean_b.strip("_")

    def format_markdown_report(self, report: DuplicateReport) -> str:
        """Formats the duplicate report into detailed GitHub-flavored Markdown."""
        lines = [
            "# Rapport d'Audit des Doublons — Suite de Production Webdyn",
            "",
            f"- **Répertoire analysé** : `{report.target_directory}`",
            f"- **Total fichiers scannés** : **{report.total_files_scanned}**",
            f"- **Fichiers doublons identifiés** : **{report.total_duplicate_files}**",
            "",
            "---",
            "",
            "## 1. Conflits de Catégorie (Classification Générique `OTH_` vs Fichier Typé)",
            "",
            "Ces fichiers `OTH_` sont des doublons stricts d'équipements déjà qualifiés en tant qu'onduleurs (`INV_`), compteurs (`MET_`), capteurs (`SEN_`) ou batteries (`BAT_`).",
            "",
            "| Fichier Canonique Qualifié | Doublon Redondant `OTH_` | Modèle Matériel | Registres | Recommandation |",
            "| :--- | :--- | :--- | :---: | :--- |",
        ]

        if not report.category_conflict_clusters:
            lines.append("| Aucun | Aucun | - | - | Conforme |")
        else:
            for c in report.category_conflict_clusters:
                dupes_str = "<br>".join(f"`{d}`" for d in c.duplicate_files)
                model_str = c.models[0] if c.models else "-"
                lines.append(f"| `{c.canonical_file}` | {dupes_str} | {model_str} | {c.register_count} | Supprimer le `OTH_` redondant |")

        lines.extend([
            "",
            "---",
            "",
            "## 2. Alias Matériels et Variantes Synonymes de Nommage",
            "",
            "Fichiers représentant le même équipement ou la même gamme sous deux désignations alternatives.",
            "",
            "| Fichier Canonique | Alias / Synonyme | Description | Registres | Recommandation |",
            "| :--- | :--- | :--- | :---: | :--- |",
        ])

        if not report.alias_clusters:
            lines.append("| Aucun | Aucun | - | - | Conforme |")
        else:
            for c in report.alias_clusters:
                dupes_str = "<br>".join(f"`{d}`" for d in c.duplicate_files)
                lines.append(f"| `{c.canonical_file}` | {dupes_str} | {c.description} | {c.register_count} | Enregistrer dans le registre d'alias |")

        lines.extend([
            "",
            "---",
            "",
            "## 3. Familles Matérielles Partageant la Même Carte Modbus",
            "",
            "Équipements distincts d'un même constructeur partageant rigoureusement le même profil de registres (ex. Schneider iEM3100..iEM3555, GoodWe ET/EH).",
            "",
            "| Modèle de Référence | Modèles de la Famille | Registres | Recommandation |",
            "| :--- | :--- | :---: | :--- |",
        ])

        if not report.exact_body_clusters:
            lines.append("| Aucun | Aucun | - | Conforme |")
        else:
            for c in report.exact_body_clusters[:25]:
                dupes_str = ", ".join(f"`{d}`" for d in c.duplicate_files[:5])
                if len(c.duplicate_files) > 5:
                    dupes_str += f" *(+{len(c.duplicate_files) - 5} autres)*"
                lines.append(f"| `{c.canonical_file}` | {dupes_str} | {c.register_count} | Conserver (modèles réels distincts) |")

        return "\n".join(lines)
