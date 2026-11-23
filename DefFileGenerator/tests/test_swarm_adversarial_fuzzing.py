"""
ADVERSARIAL STRESS & FUZZ TEST SUITE (RED-TEAM)
================================================
Apollo-Standard Adversarial, Fuzz, and Concurrency Test Suite
for Webdyn Modbus Definition Version Control & Duplicate Detection Engines.

Target Components:
1. `DefFileGenerator.version_manager.DefinitionVersionControl`
2. `DefFileGenerator.duplicate_detector.DuplicateDetector`

Apollo Invariants Verified:
- Invariant 1: Valid Input Gatekeeping (path traversal, corrupted blobs, non-existent targets).
- Invariant 3: Boundary Normalization (finities, scale boundaries, huge collections).
- Invariant 4: Honest Outcome Reporting (error propagation, tamper detection, failure reporting).
- Invariant 7: Atomic Persistence (atomic writes, crash-resilient replacements).
"""

from __future__ import annotations

import csv
import tempfile
import threading
import time
import unittest
from pathlib import Path

from DefFileGenerator.duplicate_detector import DuplicateDetector
from DefFileGenerator.version_manager import CommitInfo, DefinitionVersionControl


class TestVersionControlCASAndIntegrity(unittest.TestCase):
    """
    Adversarial verification of Content-Addressable Storage (CAS),
    cryptographic integrity, and tamper detection.
    """

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.store_dir = self.root / ".def_store"
        self.work_dir = self.root / "work"
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.vc = DefinitionVersionControl(repo_root=self.root, store_dir=self.store_dir)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def _write_csv(self, filename: str, rows: list[list[str]]) -> Path:
        p = self.work_dir / filename
        with open(p, "w", encoding="utf-8", newline="") as fp:
            writer = csv.writer(fp, delimiter=";")
            writer.writerows(rows)
        return p

    def test_tampered_cas_blob_detection(self) -> None:
        """
        Tamper with a stored CAS blob (mismatched SHA-256) and verify that
        any attempt to read, restore, or diff it triggers an explicit corruption error.
        """
        rows = [
            ["Modbus", "Inverter", "SunPower", "SPR-100", "1", "", "", "", "", "", ""],
            ["1", "3", "100", "U16", "0", "AC Power", "P_AC", "1.000000", "0.000000", "W", "4"],
        ]
        self._write_csv("INV_SunPower_SPR-100.csv", rows)
        commit = self.vc.commit(self.work_dir, message="Initial commit")

        item = commit.manifest["INV_SunPower_SPR-100.csv"]
        sha = item.sha256
        blob_path = self.store_dir / "blobs" / sha[:2] / sha[2:]
        self.assertTrue(blob_path.is_file())

        # Tamper with the blob by altering a single byte
        original_bytes = blob_path.read_bytes()
        tampered_bytes = original_bytes + b"\nTAMPERED_MALICIOUS_PAYLOAD"
        blob_path.write_bytes(tampered_bytes)

        # 1. Direct _read_blob must raise ValueError
        with self.assertRaises(ValueError) as ctx:
            self.vc._read_blob(sha)
        self.assertIn("CAS blob corruption detected", str(ctx.exception))

        # 2. get_file_content must fail honestly
        with self.assertRaises(ValueError):
            self.vc.get_file_content(commit.commit_id, "INV_SunPower_SPR-100.csv")

        # 3. restore must detect corruption and fail without writing corrupt data
        restore_dir = self.root / "restored"
        with self.assertRaises(ValueError):
            self.vc.restore(commit.commit_id, target_dir=restore_dir)

        # 4. verify_integrity audit method must flag the tampered blob
        audit = self.vc.verify_integrity()
        self.assertGreaterEqual(len(audit["corrupted_blobs"]), 1)
        self.assertIn("CAS blob corruption detected", audit["corrupted_blobs"][0])

    def test_missing_cas_blob_detection(self) -> None:
        """
        Delete a CAS blob and verify honest FileNotFoundError and integrity audit detection.
        """
        rows = [
            ["Modbus", "Meter", "Schneider", "iEM3100", "1", "", "", "", "", "", ""],
            ["1", "3", "3000", "U32", "0", "Total Energy", "E_TOT", "1.0", "0.0", "kWh", "4"],
        ]
        self._write_csv("MET_Schneider_iEM3100.csv", rows)
        commit = self.vc.commit(self.work_dir, message="Add meter")

        sha = commit.manifest["MET_Schneider_iEM3100.csv"].sha256
        blob_path = self.store_dir / "blobs" / sha[:2] / sha[2:]
        blob_path.unlink()

        # Reading or restoring missing blob raises FileNotFoundError
        with self.assertRaises(FileNotFoundError):
            self.vc._read_blob(sha)

        with self.assertRaises(FileNotFoundError):
            self.vc.restore(commit.commit_id, target_dir=self.root / "restore")

        # Integrity audit flags missing blob
        audit = self.vc.verify_integrity()
        self.assertEqual(len(audit["missing_blobs"]), 1)
        self.assertIn("MET_Schneider_iEM3100.csv", audit["missing_blobs"][0])

    def test_cas_blob_deduplication(self) -> None:
        """
        Verify that identical file contents are deduplicated into a single CAS blob.
        """
        rows = [
            ["Modbus", "Inverter", "SolarMax", "SM10", "1", "", "", "", "", "", ""],
            ["1", "3", "10", "U16", "0", "Freq", "FREQ", "0.01", "0.0", "Hz", "4"],
        ]
        self._write_csv("INV_SolarMax_SM10_A.csv", rows)
        self._write_csv("INV_SolarMax_SM10_B.csv", rows)

        commit = self.vc.commit(self.work_dir, message="Two identical files")
        self.assertEqual(commit.files_count, 2)

        sha_a = commit.manifest["INV_SolarMax_SM10_A.csv"].sha256
        sha_b = commit.manifest["INV_SolarMax_SM10_B.csv"].sha256
        self.assertEqual(sha_a, sha_b)

        # Count stored blob files on disk
        stored_blobs = [p for p in self.vc.blobs_dir.rglob("*") if p.is_file()]
        self.assertEqual(len(stored_blobs), 1)


class TestVersionControlEdgeCommitsAndLineage(unittest.TestCase):
    """
    Adversarial tests for boundary commit states, empty commits,
    non-existent commits, and corrupted index states.
    """

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.store_dir = self.root / ".def_store"
        self.work_dir = self.root / "work"
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.vc = DefinitionVersionControl(repo_root=self.root, store_dir=self.store_dir)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_empty_commit_with_zero_files(self) -> None:
        """
        Verify committing an empty directory with 0 files succeeds honestly
        without crashing and records 0 files, 0 registers.
        """
        commit = self.vc.commit(self.work_dir, message="Empty initial commit")
        self.assertEqual(commit.files_count, 0)
        self.assertEqual(commit.total_registers, 0)
        self.assertEqual(commit.total_tags, 0)
        self.assertEqual(commit.manifest, {})
        self.assertEqual(self.vc.get_head(), commit.commit_id)

        # Restoring empty commit produces 0 restored files
        restored = self.vc.restore(commit.commit_id, target_dir=self.root / "target")
        self.assertEqual(len(restored), 0)

    def test_nonexistent_commit_rejection(self) -> None:
        """
        Verify querying non-existent commit IDs raises KeyError across all entrypoints.
        """
        bad_id = "0000000000000000"
        with self.assertRaises(KeyError):
            self.vc.get_commit(bad_id)

        with self.assertRaises(KeyError):
            self.vc.restore(bad_id, target_dir=self.root / "target")

        with self.assertRaises(KeyError):
            self.vc.diff(bad_id)

        with self.assertRaises(KeyError):
            self.vc.get_file_content(bad_id, "file.csv")

        with self.assertRaises(KeyError):
            self.vc.create_tag("tag1", bad_id)

    def test_nonexistent_source_directory(self) -> None:
        """
        Verify committing a non-existent directory raises NotADirectoryError.
        """
        bad_dir = self.root / "does_not_exist"
        with self.assertRaises(NotADirectoryError):
            self.vc.commit(bad_dir, message="Will fail")

    def test_restore_nonexistent_file_in_valid_commit(self) -> None:
        """
        Verify attempting to restore a file that does not exist in the manifest raises FileNotFoundError.
        """
        p = self.work_dir / "INV_Sample.csv"
        p.write_text("Modbus;INV;A;B;1\n1;3;1;U16;0;R;T;1;0;W;4\n", encoding="utf-8")
        c = self.vc.commit(self.work_dir, message="Valid commit")

        with self.assertRaises(FileNotFoundError):
            self.vc.restore(c.commit_id, target_dir=self.root / "out", file_name="NON_EXISTENT.csv")

    def test_tag_on_empty_repository_raises_value_error(self) -> None:
        """
        Verify creating a tag when HEAD is None and no commit ID is provided raises ValueError.
        """
        with self.assertRaises(ValueError) as ctx:
            self.vc.create_tag("v1.0")
        self.assertIn("repository is empty", str(ctx.exception))


class TestVersionControlPathTraversalSecurity(unittest.TestCase):
    """
    Hostile boundary testing for directory traversal, escaped paths,
    and symlink/injection attacks.
    """

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.store_dir = self.root / ".def_store"
        self.work_dir = self.root / "work"
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.vc = DefinitionVersionControl(repo_root=self.root, store_dir=self.store_dir)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_commit_path_traversal_rejection(self) -> None:
        """
        Reject path traversal patterns in files parameter for commit().
        """
        malicious_files = [
            "../../escape.csv",
            "../secret.csv",
            "..\\windows_escape.csv",
            "/etc/passwd.csv",
            "subdir/nested.csv",
        ]
        for bad_f in malicious_files:
            with self.subTest(file=bad_f):
                with self.assertRaises(ValueError) as ctx:
                    self.vc.commit(self.work_dir, message="Hostile", files=[bad_f])
                self.assertIn("Path traversal detected", str(ctx.exception))

    def test_restore_path_traversal_rejection(self) -> None:
        """
        Reject path traversal attempts in restore() file_name parameter.
        """
        p = self.work_dir / "INV_Valid.csv"
        p.write_text("Modbus;INV;Vendor;Model;1\n1;3;1;U16;0;R;T;1;0;;4\n", encoding="utf-8")
        c = self.vc.commit(self.work_dir, message="Clean commit")

        malicious_restores = [
            "../../etc/passwd",
            "..\\escaped.csv",
            "sub/folder.csv",
        ]
        for bad_restore in malicious_restores:
            with self.subTest(restore=bad_restore):
                with self.assertRaises((ValueError, FileNotFoundError)):
                    self.vc.restore(
                        c.commit_id, target_dir=self.root / "out", file_name=bad_restore
                    )

    def test_tag_name_path_traversal_rejection(self) -> None:
        """
        Reject tag names containing path separators or traversal characters.
        """
        p = self.work_dir / "INV_Valid.csv"
        p.write_text("Modbus;INV;Vendor;Model;1\n1;3;1;U16;0;R;T;1;0;;4\n", encoding="utf-8")
        c = self.vc.commit(self.work_dir, message="Clean commit")

        malicious_tags = [
            "../../bad_tag",
            "..\\windows_tag",
            "tag/with/slash",
            "tag\\backslash",
            "",
            "   ",
            "tag with spaces",
            "tag$name",
        ]
        for bad_tag in malicious_tags:
            with self.subTest(tag=bad_tag):
                with self.assertRaises(ValueError):
                    self.vc.create_tag(bad_tag, commit_id=c.commit_id)

    def test_file_name_traversal_in_queries(self) -> None:
        """
        Reject path traversal in log(), diff(), and get_file_content().
        """
        p = self.work_dir / "INV_Valid.csv"
        p.write_text("Modbus;INV;Vendor;Model;1\n1;3;1;U16;0;R;T;1;0;;4\n", encoding="utf-8")
        c = self.vc.commit(self.work_dir, message="Clean commit")

        with self.assertRaises(ValueError):
            self.vc.log(file_name="../../escape.csv")

        with self.assertRaises(ValueError):
            self.vc.diff(c.commit_id, file_name="../../escape.csv")

        with self.assertRaises(ValueError):
            self.vc.get_file_content(c.commit_id, file_name="../../escape.csv")


class TestVersionControlUnicodeAndMalformedCSV(unittest.TestCase):
    """
    Fuzzing definition parser and commit engine with multi-byte Unicode,
    exotic character sets, and corrupted/malformed CSV content.
    """

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.store_dir = self.root / ".def_store"
        self.work_dir = self.root / "work"
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.vc = DefinitionVersionControl(repo_root=self.root, store_dir=self.store_dir)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_multilingual_unicode_support(self) -> None:
        """
        Verify Chinese, Japanese, European accents, and Emoji in definitions
        survive commit, manifest parsing, diff, and byte-exact restoration.
        """
        filename = "INV_Huawei_华为SUN2000.csv"
        p = self.work_dir / filename
        content = (
            "Modbus;Inverter;华为技术有限公司;SUN2000-100KTL-M1 ⚡;1;;;;;;\n"
            "1;3;30000;U16;0;输入电压 (Input Voltage);U_DC;0.100000;0.000000;V;4\n"
            "2;3;30001;I32;0;Température ambiante (°C);TEMP_AMB;0.010000;0.000000;°C;4\n"
            "3;3;30003;U32;0;累计发电量 (Solar Yield ☀️);E_TOT;1.000000;0.000000;kWh;4\n"
        )
        p.write_text(content, encoding="utf-8")

        commit = self.vc.commit(
            self.work_dir, message="Commit multilingual definition", author="Éléonore François"
        )
        self.assertEqual(commit.files_count, 1)
        self.assertEqual(commit.total_registers, 3)
        self.assertEqual(commit.total_tags, 3)

        manifest_item = commit.manifest[filename]
        self.assertEqual(manifest_item.manufacturer, "华为技术有限公司")
        self.assertIn("SUN2000-100KTL-M1 ⚡", manifest_item.model)

        # Restore and verify byte-for-byte exact equality
        restore_dir = self.root / "restore_unicode"
        restored_files = self.vc.restore(commit.commit_id, target_dir=restore_dir)
        self.assertEqual(len(restored_files), 1)
        restored_bytes = restored_files[0].read_bytes()
        self.assertEqual(restored_bytes, p.read_bytes())

    def test_malformed_csv_fuzzing(self) -> None:
        """
        Fuzz the metadata parser and diff engine with malformed CSV inputs:
        - 0 bytes empty file
        - 1 single byte
        - Header with 1 column
        - Missing columns in register rows (< 11 columns)
        - Overflow columns (> 30 columns)
        - Null bytes and random binary data
        """
        malformed_samples = {
            "empty.csv": b"",
            "single_byte.csv": b";",
            "short_header.csv": b"Modbus\n",
            "truncated_cols.csv": b"Modbus;INV;Vendor;Model;1\n1;3;100\n2;3;101;U16\n",
            "overflow_cols.csv": b"Modbus;INV;Vendor;Model;1;;;;;;;;;;;;;;;;;;;;;\n"
            b"1;3;100;U16;0;Reg;Tag;1;0;W;4;extra1;extra2;extra3\n",
            "binary_with_null.csv": b"Modbus;INV;Vendor;Model;1\n\x00\x01\x02\xff;3;100\n",
        }

        for fname, raw_data in malformed_samples.items():
            (self.work_dir / fname).write_bytes(raw_data)

        # Committing malformed files must NOT crash
        commit = self.vc.commit(self.work_dir, message="Malformed files commit")
        self.assertEqual(commit.files_count, len(malformed_samples))

        # Check diff against itself must be clean
        diff_res = self.vc.diff(commit.commit_id, commit.commit_id)
        self.assertEqual(len(diff_res["modified_files"]), 0)


class TestVersionControlSemanticDiffStress(unittest.TestCase):
    """
    Stress testing semantic Modbus register diffing with massive register tables (>10,000 registers).
    """

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.store_dir = self.root / ".def_store"
        self.work_dir = self.root / "work"
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.vc = DefinitionVersionControl(repo_root=self.root, store_dir=self.store_dir)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_massive_semantic_diff_12000_registers(self) -> None:
        """
        Verify semantic diffing performance and accuracy on a massive dataset:
        Commit 1: 10,000 registers.
        Commit 2:
          - 2,000 registers added (total 12,000)
          - 1,500 registers removed
          - 3,000 registers modified (type, tag, scale changes)
          - 5,500 registers unchanged
        Diff must complete in < 3.0 seconds with exact delta statistics.
        """
        fname = "INV_Giant_SolarPlant.csv"
        file_path = self.work_dir / fname

        # --- Generate Commit 1 (10,000 registers: addresses 1000 to 10999) ---
        header = ["Modbus", "Inverter", "GiantMfg", "MegaInverter_V1", "1", "", "", "", "", "", ""]
        rows_v1 = [header]
        for addr in range(1000, 11000):
            idx = addr - 999
            rows_v1.append(
                [
                    str(idx),
                    "3",
                    str(addr),
                    "U16",
                    "0",
                    f"Register_{addr}",
                    f"TAG_{addr}",
                    "1.000000",
                    "0.000000",
                    "W",
                    "4",
                ]
            )

        with open(file_path, "w", encoding="utf-8", newline="") as fp:
            writer = csv.writer(fp, delimiter=";")
            writer.writerows(rows_v1)

        c1 = self.vc.commit(self.work_dir, message="V1: 10,000 registers")
        self.assertEqual(c1.total_registers, 10000)

        # --- Generate Commit 2 ---
        # 1. Keep 5,500 registers (addr 1000 to 6499) unchanged
        # 2. Modify 3,000 registers (addr 6500 to 9499) with modified data type and tag
        # 3. Remove 1,500 registers (addr 9500 to 10999) - omitted from rows_v2
        # 4. Add 2,000 new registers (addr 20000 to 21999)
        header_v2 = [
            "Modbus",
            "Inverter",
            "GiantMfg",
            "MegaInverter_V2",
            "1",
            "",
            "",
            "",
            "",
            "",
            "",
        ]
        rows_v2 = [header_v2]

        cur_idx = 1
        for addr in range(1000, 6500):
            rows_v2.append(
                [
                    str(cur_idx),
                    "3",
                    str(addr),
                    "U16",
                    "0",
                    f"Register_{addr}",
                    f"TAG_{addr}",
                    "1.000000",
                    "0.000000",
                    "W",
                    "4",
                ]
            )
            cur_idx += 1

        for addr in range(6500, 9500):
            # Change U16 -> I32, and alter tag
            rows_v2.append(
                [
                    str(cur_idx),
                    "3",
                    str(addr),
                    "I32",
                    "0",
                    f"Register_{addr}",
                    f"TAG_{addr}_MOD",
                    "0.100000",
                    "0.000000",
                    "kW",
                    "4",
                ]
            )
            cur_idx += 1

        for addr in range(20000, 22000):
            rows_v2.append(
                [
                    str(cur_idx),
                    "3",
                    str(addr),
                    "U32",
                    "0",
                    f"NewReg_{addr}",
                    f"NEW_TAG_{addr}",
                    "1.000000",
                    "0.000000",
                    "V",
                    "4",
                ]
            )
            cur_idx += 1

        with open(file_path, "w", encoding="utf-8", newline="") as fp:
            writer = csv.writer(fp, delimiter=";")
            writer.writerows(rows_v2)

        c2 = self.vc.commit(self.work_dir, message="V2: modified 3000, removed 1500, added 2000")
        self.assertEqual(c2.total_registers, 5500 + 3000 + 2000)  # 10500 total

        # --- Compute Diff ---
        t0 = time.perf_counter()
        diff_res = self.vc.diff(c2.commit_id, c1.commit_id)
        elapsed = time.perf_counter() - t0

        self.assertLess(elapsed, 3.0, f"Diff took {elapsed:.2f}s, expected < 3.0s")
        self.assertIn(fname, diff_res["modified_files"])

        f_diff = diff_res["file_diffs"][fname]
        self.assertTrue(f_diff["header_changed"])
        self.assertEqual(f_diff["added_registers_count"], 2000)
        self.assertEqual(f_diff["removed_registers_count"], 1500)
        self.assertEqual(f_diff["modified_registers_count"], 3000)

    def test_diff_with_hex_and_exotic_addresses(self) -> None:
        """
        Verify register address sorting and diffing with hex ("0x1000"),
        hyphenated ("100-2"), and alphanumeric keys without crashing.
        """
        fname = "INV_HexAddress.csv"
        file_path = self.work_dir / fname

        rows_a = [
            ["Modbus", "INV", "Brand", "Model", "1", "", "", "", "", "", ""],
            ["1", "3", "0x1000", "U16", "0", "RegA", "TAG_A", "1", "0", "V", "4"],
            ["2", "3", "NaN_ADDR", "U16", "0", "RegB", "TAG_B", "1", "0", "V", "4"],
        ]
        with open(file_path, "w", encoding="utf-8", newline="") as fp:
            csv.writer(fp, delimiter=";").writerows(rows_a)
        c1 = self.vc.commit(self.work_dir, message="c1 hex")

        rows_b = [
            ["Modbus", "INV", "Brand", "Model", "1", "", "", "", "", "", ""],
            ["1", "3", "0x1000", "I32", "0", "RegA", "TAG_A_NEW", "1", "0", "V", "4"],
            ["2", "3", "0x2000", "U16", "0", "RegC", "TAG_C", "1", "0", "V", "4"],
        ]
        with open(file_path, "w", encoding="utf-8", newline="") as fp:
            csv.writer(fp, delimiter=";").writerows(rows_b)
        c2 = self.vc.commit(self.work_dir, message="c2 hex update")

        diff_res = self.vc.diff(c2.commit_id, c1.commit_id)
        f_diff = diff_res["file_diffs"][fname]
        self.assertEqual(f_diff["added_registers_count"], 1)
        self.assertEqual(f_diff["removed_registers_count"], 1)
        self.assertEqual(f_diff["modified_registers_count"], 1)


class TestVersionControlConcurrencyAndLockSafety(unittest.TestCase):
    """
    Multithreaded stress testing to verify lock safety, atomic commits,
    and index file integrity under concurrent load.
    """

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.store_dir = self.root / ".def_store"
        self.vc = DefinitionVersionControl(repo_root=self.root, store_dir=self.store_dir)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_concurrent_commits_thread_safety(self) -> None:
        """
        Execute 12 simultaneous threads each committing a separate file to the repository.
        Verify that all commits land atomically, index.json remains uncorrupted,
        and no commits are lost.
        """
        num_threads = 12
        thread_dirs: list[Path] = []
        for i in range(num_threads):
            tdir = self.root / f"worker_{i}"
            tdir.mkdir(parents=True, exist_ok=True)
            fname = f"INV_Worker_{i:02d}.csv"
            (tdir / fname).write_text(
                f"Modbus;INV;Vendor;WorkerModel_{i};1\n1;3;{i * 100};U16;0;Reg_{i};TAG_{i};1;0;W;4\n",
                encoding="utf-8",
            )
            thread_dirs.append(tdir)

        results: list[CommitInfo] = []
        errors: list[Exception] = []

        def worker_commit(worker_idx: int) -> None:
            try:
                commit = self.vc.commit(
                    source_dir=thread_dirs[worker_idx],
                    message=f"Commit from worker {worker_idx}",
                    author=f"Worker-{worker_idx}",
                )
                results.append(commit)
            except Exception as e:
                errors.append(e)

        threads = [threading.Thread(target=worker_commit, args=(i,)) for i in range(num_threads)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        self.assertEqual(len(errors), 0, f"Encountered concurrency errors: {errors}")
        self.assertEqual(len(results), num_threads)

        # Verify index.json integrity
        all_commits = self.vc.log(limit=100)
        self.assertEqual(len(all_commits), num_threads)

        # Integrity audit must pass cleanly
        audit = self.vc.verify_integrity()
        self.assertEqual(len(audit["corrupted_blobs"]), 0)
        self.assertEqual(len(audit["missing_blobs"]), 0)
        self.assertEqual(audit["commits_checked"], num_threads)


class TestDuplicateDetectorAdversarialFuzzing(unittest.TestCase):
    """
    Adversarial and fuzz testing for DuplicateDetector:
    - Malformed CSV, missing headers, inconsistent columns.
    - NaN, Infinity, negative values.
    - Empty files, single-line files, non-CSV files.
    - False-positive avoidance on zero-register files.
    - Hardware alias extraction edge cases.
    """

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.work_dir = Path(self.tmp_dir.name)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def _write_file(self, filename: str, content: str | bytes) -> Path:
        p = self.work_dir / filename
        if isinstance(content, str):
            p.write_text(content, encoding="utf-8")
        else:
            p.write_bytes(content)
        return p

    def test_empty_and_single_line_files_no_false_body_clusters(self) -> None:
        """
        Zero-register files (empty files, header-only single lines) must NOT
        be grouped into exact_body hardware families.
        """
        # Two empty files (0 bytes) -> Byte identical exact_sha256 duplicates
        self._write_file("INV_Empty1.csv", "")
        self._write_file("INV_Empty2.csv", "")

        # Two different single-line files with 0 registers
        self._write_file("INV_BrandA_Solar.csv", "Modbus;INV;BrandA;Solar;1;;;;;;\n")
        self._write_file("MET_BrandB_Meter.csv", "Modbus;MET;BrandB;Meter;1;;;;;;\n")

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        # Empty files share 0 bytes SHA-256
        self.assertEqual(len(report.exact_hash_clusters), 1)
        self.assertEqual(report.exact_hash_clusters[0].canonical_file, "INV_Empty1.csv")

        # But single-line files with 0 registers must NOT form exact_body clusters
        self.assertEqual(len(report.exact_body_clusters), 0)
        self.assertEqual(len(report.category_conflict_clusters), 0)
        self.assertEqual(len(report.alias_clusters), 0)

    def test_non_csv_and_binary_garbage_ignored(self) -> None:
        """
        Non-CSV files (.txt, .json, .bin) are ignored; corrupted binary files with .csv
        extension are processed without raising unhandled exceptions.
        """
        self._write_file("notes.txt", "Some random notes")
        self._write_file("data.json", '{"key": "value"}')
        self._write_file("backup.bin", b"\x00\x01\x02\x03\x04")

        # Binary file with .csv extension containing null bytes
        self._write_file("INV_Corrupt.csv", b"\x00\xff\xfe\x00Modbus\x00;Corrupted;\x00")

        # Valid file
        self._write_file(
            "INV_Valid.csv",
            "Modbus;INV;Brand;Model;1\n1;3;100;U16;0;Power;P;1;0;W;4\n",
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        self.assertEqual(report.total_files_scanned, 2)  # INV_Corrupt.csv and INV_Valid.csv
        self.assertEqual(report.total_duplicate_files, 0)

    def test_malformed_csv_inconsistent_columns_nan_infinity(self) -> None:
        """
        Definitions with missing headers, non-numeric register addresses,
        NaN/Infinity values, and inconsistent columns must be handled robustly.
        """
        # File A: Addresses are NaN and Infinity
        content_a = (
            "Modbus;Inverter;WeirdMfg;WeirdModel;1\n"
            "1;3;NaN;U16;0;Voltage;U;NaN;0;V;4\n"
            "2;3;Infinity;I32;0;Current;I;1.0;Infinity;A;4\n"
        )
        self._write_file("INV_WeirdMfg_WeirdModel.csv", content_a)

        # File B: Duplicate register table of File A under OTH_ prefix
        content_b = (
            "Modbus;Other;WeirdMfg;WeirdModel;1\n"
            "1;3;NaN;U16;0;Voltage;U;NaN;0;V;4\n"
            "2;3;Infinity;I32;0;Current;I;1.0;Infinity;A;4\n"
        )
        self._write_file("OTH_WeirdMfg_WeirdModel.csv", content_b)

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        # Must correctly identify category conflict despite NaN / Infinity in cells
        self.assertEqual(len(report.category_conflict_clusters), 1)
        cluster = report.category_conflict_clusters[0]
        self.assertEqual(cluster.canonical_file, "INV_WeirdMfg_WeirdModel.csv")
        self.assertEqual(cluster.duplicate_files, ["OTH_WeirdMfg_WeirdModel.csv"])

    def test_alias_extraction_boundary_cases(self) -> None:
        """
        Verify alias detection logic:
        1. True alias variations (e.g. HT1500 vs HT_1500V, GT_Series vs GT-SERIES) match.
        2. Names that strip down to empty strings (e.g. INV_Generic.csv vs INV_Series.csv) do NOT false-match.
        """
        # Shared register body
        body = "1;3;10;U16;0;Reg;Tag;1;0;W;4\n"

        # Case 1: True alias
        self._write_file("INV_Goodwe_HT1500.csv", f"Modbus;INV;Goodwe;HT1500;1\n{body}")
        self._write_file("INV_Goodwe_HT_1500V.csv", f"Modbus;INV;Goodwe;HT_1500V;1\n{body}")

        # Case 2: Distinct names whose stems both happen to contain stripped tokens ("generic", "series")
        # but are distinct models:
        self._write_file(
            "INV_Generic.csv", f"Modbus;INV;Brand;Generic;1\n{body}2;3;11;U16;0;R2;T2;1;0;;4\n"
        )
        self._write_file(
            "INV_Series.csv", f"Modbus;INV;Brand;Series;1\n{body}2;3;11;U16;0;R2;T2;1;0;;4\n"
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        # HT1500 vs HT_1500V must be detected as an alias cluster
        alias_names = [c.canonical_file for c in report.alias_clusters]
        self.assertIn("INV_Goodwe_HT1500.csv", alias_names)

        # INV_Generic vs INV_Series must NOT be grouped into an alias cluster
        alias_canonicals = {c.canonical_file for c in report.alias_clusters}
        alias_dupes = {d for c in report.alias_clusters for d in c.duplicate_files}
        self.assertNotIn("INV_Generic.csv", alias_canonicals | alias_dupes)

    def test_hardware_family_grouping(self) -> None:
        """
        Verify distinct hardware models sharing standard register tables
        (e.g., Schneider iEM3150 and iEM3155) are classified into exact_body_clusters.
        """
        body = (
            "1;3;3000;U32;0;Active Energy;E_ACT;1.0;0.0;kWh;4\n"
            "2;3;3002;U32;0;Reactive Energy;E_REA;1.0;0.0;kVARh;4\n"
        )
        self._write_file("MET_Schneider_iEM3150.csv", f"Modbus;MET;Schneider;iEM3150;1\n{body}")
        self._write_file("MET_Schneider_iEM3155.csv", f"Modbus;MET;Schneider;iEM3155;1\n{body}")

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        self.assertEqual(len(report.exact_body_clusters), 1)
        cluster = report.exact_body_clusters[0]
        self.assertEqual(cluster.cluster_type, "exact_body")
        self.assertEqual(cluster.register_count, 2)
        self.assertEqual(cluster.action_recommendation, "keep_hardware_family_models")

    def test_markdown_report_formatting_adversarial(self) -> None:
        """
        Generate markdown report under empty and complex conditions, verifying valid output.
        """
        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()
        md = detector.format_markdown_report(report)

        self.assertIn("# Rapport d'Audit des Doublons", md)
        self.assertIn("Total fichiers scannés", md)
        self.assertIn("Aucun", md)


class TestApolloInvariantsEnforcement(unittest.TestCase):
    """
    Direct verification of Apollo Non-Negotiable Engineering Invariants:
    - Invariant 1: Valid Input Gatekeeping
    - Invariant 4: Honest Outcome Reporting
    - Invariant 7: Atomic Persistence
    """

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.store_dir = self.root / ".def_store"
        self.work_dir = self.root / "work"
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.vc = DefinitionVersionControl(repo_root=self.root, store_dir=self.store_dir)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_invariant_1_valid_input_gatekeeping(self) -> None:
        """
        Invalid inputs (empty commit IDs, None paths, traversal tags) must be rejected
        at the boundary before any mutation occurs.
        """
        with self.assertRaises(ValueError):
            self.vc.create_tag("")

        with self.assertRaises(ValueError):
            self.vc.create_tag("invalid/tag/path")

        with self.assertRaises(KeyError):
            self.vc.get_commit("invalid_commit_id")

    def test_invariant_7_atomic_persistence(self) -> None:
        """
        Verify writes use atomic temp-file replace patterns with zero leftover temporary files.
        """
        p = self.work_dir / "INV_Sample.csv"
        p.write_text("Modbus;INV;Vendor;Model;1\n1;3;100;U16;0;Reg;Tag;1;0;W;4\n", encoding="utf-8")
        c = self.vc.commit(self.work_dir, message="Initial commit")

        # Verify index.json and commit JSON exist without leftover .tmp files
        tmp_files = list(self.store_dir.rglob("*.tmp"))
        self.assertEqual(len(tmp_files), 0, f"Found leftover temporary files: {tmp_files}")

        # Restore must not leave .restore_tmp files
        target = self.root / "restored"
        self.vc.restore(c.commit_id, target_dir=target)
        restore_tmps = list(target.rglob("*.restore_tmp"))
        self.assertEqual(
            len(restore_tmps), 0, f"Found leftover restore temporary files: {restore_tmps}"
        )


if __name__ == "__main__":
    unittest.main()
