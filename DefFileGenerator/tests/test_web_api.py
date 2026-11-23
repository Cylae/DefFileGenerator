"""
Test suite for Web UI assets and FastAPI endpoints for Versions & Commits
and Duplicates & Quality Dashboard.
"""

import tempfile
import unittest

from fastapi.testclient import TestClient

from web.app import app


class TestWebApiAndStaticAssets(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)

    def test_root_index_contains_new_tabs(self) -> None:
        """Verifies that index.html contains the new tabs, panes, and modal."""
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        content = response.text

        # Navigation tabs
        self.assertIn('id="tabBtnVersions"', content)
        self.assertIn('data-tab="tab-versions"', content)
        self.assertIn("Versions & Commits", content)

        self.assertIn('id="tabBtnDuplicates"', content)
        self.assertIn('data-tab="tab-duplicates"', content)
        self.assertIn("Duplicates & Quality", content)

        # Tab Panes
        self.assertIn('id="tab-versions"', content)
        self.assertIn('id="tab-duplicates"', content)

        # Restore modal
        self.assertIn('id="restoreModal"', content)
        self.assertIn('id="confirmRestoreBtn"', content)
        self.assertIn('id="restoreTargetDir"', content)
        self.assertIn('id="restorePreviewList"', content)

        # 5 Pillars cards
        self.assertIn("1. UTF-8 without BOM", content)
        self.assertIn("2. CRLF Line Endings", content)
        self.assertIn("3. Strict 11 Columns", content)
        self.assertIn("4. Sequential Index", content)
        self.assertIn("5. Qualified Types & Tags", content)

        # Duplicate section tables
        self.assertIn('id="categoryConflictTable"', content)
        self.assertIn('id="aliasTable"', content)
        self.assertIn('id="hardwareFamilyTable"', content)

    def test_static_app_js_contains_handlers(self) -> None:
        """Verifies that /static/app.js serves the logic for both new tabs."""
        response = self.client.get("/static/app.js")
        self.assertEqual(response.status_code, 200)
        js = response.text

        # Tab bindings
        self.assertIn("tabBtnVersions", js)
        self.assertIn("tabBtnDuplicates", js)

        # Versions logic
        self.assertIn("loadVersions", js)
        self.assertIn("inspectCommit", js)
        self.assertIn("renderCommitDiff", js)
        self.assertIn("renderCommitManifest", js)
        self.assertIn("openRestoreModal", js)

        # Duplicates logic
        self.assertIn("loadDuplicates", js)
        self.assertIn("renderDuplicates", js)
        self.assertIn("renderCategoryConflicts", js)
        self.assertIn("renderAliasClusters", js)
        self.assertIn("renderHardwareFamilies", js)

    def test_api_versions_log(self) -> None:
        """Verifies GET /api/versions returns commit list."""
        response = self.client.get("/api/versions?limit=10")
        self.assertEqual(response.status_code, 200)
        commits = response.json()
        self.assertIsInstance(commits, list)
        if commits:
            commit = commits[0]
            self.assertIn("commit_id", commit)
            self.assertIn("author", commit)
            self.assertIn("timestamp", commit)
            self.assertIn("message", commit)
            self.assertIn("files_count", commit)
            self.assertTrue("total_registers" in commit or "changes" in commit)

    def test_api_versions_commit_details(self) -> None:
        """Verifies GET /api/versions/{commit_id} returns metadata and manifest."""
        history = self.client.get("/api/versions?limit=1").json()
        if not history:
            self.skipTest("No commits in history to test details")

        commit_id = history[0]["commit_id"]
        response = self.client.get(f"/api/versions/{commit_id}")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["commit_id"], commit_id)
        self.assertIn("manifest", data)
        self.assertIsInstance(data["manifest"], dict)

    def test_api_versions_diff(self) -> None:
        """Verifies GET /api/versions/{commit_id}/diff returns semantic diff."""
        history = self.client.get("/api/versions?limit=1").json()
        if not history:
            self.skipTest("No commits in history to test diff")

        commit_id = history[0]["commit_id"]
        response = self.client.get(f"/api/versions/{commit_id}/diff")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn("commit_a", data)
        self.assertIn("added_files", data)
        self.assertIn("modified_files", data)
        self.assertIn("removed_files", data)
        self.assertIn("file_diffs", data)

    def test_api_versions_not_found(self) -> None:
        """Verifies 404 on invalid commit_id."""
        response = self.client.get("/api/versions/0000000000000000")
        self.assertEqual(response.status_code, 404)

    def test_api_versions_restore_to_temp_dir(self) -> None:
        """Verifies restoring definitions from a commit to a target directory."""
        history = self.client.get("/api/versions?limit=1").json()
        if not history:
            self.skipTest("No commits in history to test restore")

        commit_id = history[0]["commit_id"]
        with tempfile.TemporaryDirectory() as tmp_dir:
            payload = {
                "commit_id": commit_id,
                "target_dir": str(tmp_dir),
                "file_name": None,
            }
            response = self.client.post("/api/versions/restore", json=payload)
            self.assertEqual(response.status_code, 200)
            data = response.json()
            self.assertTrue(data["success"])
            self.assertEqual(data["commit_id"], commit_id)
            self.assertGreater(data["restored_count"], 0)

    def test_api_duplicates_report(self) -> None:
        """Verifies GET /api/duplicates returns complete report with clusters."""
        response = self.client.get("/api/duplicates?path=production")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn("target_directory", data)
        self.assertIn("total_files_scanned", data)
        self.assertIn("total_duplicate_files", data)
        self.assertIn("category_conflict_clusters", data)
        self.assertIn("alias_clusters", data)
        self.assertIn("exact_body_clusters", data)
        self.assertGreater(data["total_files_scanned"], 0)


if __name__ == "__main__":
    unittest.main()
