"""
Unit tests for Milestone B.1: DocumentFetcher security (anti-SSRF) and IdentityResolver.
"""

import unittest
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from DefFileGenerator.discovery.document_fetcher import (
    DocumentFetcher,
    FetchOutcome,
    is_ip_allowed,
    validate_target_url,
)
from DefFileGenerator.discovery.identity_resolver import IdentityResolver


class TestDocumentFetcherSecurity(unittest.TestCase):
    """Adversarial and boundary security checks for DocumentFetcher."""

    def test_ip_allowed_filters_private_and_dangerous_ranges(self):
        # Disallowed / dangerous ranges
        self.assertFalse(is_ip_allowed("127.0.0.1"), "Loopback IPv4 must be blocked")
        self.assertFalse(is_ip_allowed("127.0.0.2"), "Loopback IPv4 range must be blocked")
        self.assertFalse(is_ip_allowed("::1"), "Loopback IPv6 must be blocked")
        self.assertFalse(is_ip_allowed("10.0.0.1"), "RFC 1918 Class A must be blocked")
        self.assertFalse(is_ip_allowed("172.16.0.1"), "RFC 1918 Class B must be blocked")
        self.assertFalse(is_ip_allowed("192.168.1.1"), "RFC 1918 Class C must be blocked")
        self.assertFalse(is_ip_allowed("169.254.169.254"), "Cloud metadata IP must be blocked")
        self.assertFalse(is_ip_allowed("169.254.1.1"), "Link-local IPv4 must be blocked")
        self.assertFalse(is_ip_allowed("fe80::1"), "Link-local IPv6 must be blocked")
        self.assertFalse(is_ip_allowed("224.0.0.1"), "Multicast must be blocked")
        self.assertFalse(is_ip_allowed("invalid-ip"), "Invalid IP strings must be rejected")

        # Allowed public routable IPs
        self.assertTrue(is_ip_allowed("8.8.8.8"), "Public IP must be allowed")
        self.assertTrue(is_ip_allowed("1.1.1.1"), "Public IP must be allowed")
        self.assertTrue(is_ip_allowed("2606:4700:4700::1111"), "Public IPv6 must be allowed")

    def test_validate_target_url_schemes_and_hosts(self):
        # Forbidden schemes
        valid, reason, _ = validate_target_url("ftp://example.com/doc.pdf")
        self.assertFalse(valid)
        self.assertIn("scheme", reason)

        valid, reason, _ = validate_target_url("file:///etc/passwd")
        self.assertFalse(valid)

        # Loopback hostnames
        valid, reason, _ = validate_target_url("http://localhost:8080/secret")
        self.assertFalse(valid)
        self.assertIn("loopback", reason)

        valid, reason, _ = validate_target_url("http://127.0.0.1/doc.pdf")
        self.assertFalse(valid)

        valid, reason, _ = validate_target_url("http://169.254.169.254/latest/meta-data/")
        self.assertFalse(valid)

    @pytest.mark.anyio
    async def test_fetcher_blocks_ssrf_preflight(self):
        fetcher = DocumentFetcher()
        result = await fetcher.fetch("http://127.0.0.1:8000/internal-api")
        self.assertEqual(result.outcome, FetchOutcome.BLOCKED_SSRF)
        self.assertIn("SSRF guard violation", result.error_message or "")
        self.assertIsNone(result.file_bytes)

    @pytest.mark.anyio
    async def test_fetcher_blocks_cloud_metadata(self):
        fetcher = DocumentFetcher()
        result = await fetcher.fetch("http://169.254.169.254/latest/meta-data/")
        self.assertEqual(result.outcome, FetchOutcome.BLOCKED_SSRF)
        self.assertIsNone(result.file_bytes)

    @pytest.mark.anyio
    async def test_fetcher_content_length_overspec_abort(self):
        fetcher = DocumentFetcher(max_bytes=1000)

        mock_resp = MagicMock()
        mock_resp.is_redirect = False
        mock_resp.status_code = 200
        mock_resp.headers = {"content-length": "5000"}

        with patch("DefFileGenerator.discovery.document_fetcher.validate_target_url", return_value=(True, "", ["93.184.216.34"])):
            with patch("httpx.AsyncClient") as mock_client_cls:
                mock_client = MagicMock()
                mock_client.__aenter__ = AsyncMock(return_value=mock_client)
                mock_client.__aexit__ = AsyncMock(return_value=None)
                mock_stream_ctx = MagicMock()
                mock_stream_ctx.__aenter__ = AsyncMock(return_value=mock_resp)
                mock_stream_ctx.__aexit__ = AsyncMock(return_value=None)
                mock_client.stream.return_value = mock_stream_ctx
                mock_client_cls.return_value = mock_client

                result = await fetcher.fetch("https://example.com/oversized.pdf")
                self.assertEqual(result.outcome, FetchOutcome.OVERSPEC_BYTES)
                self.assertIn("exceeds max allowed", result.error_message or "")


class TestIdentityResolver(unittest.TestCase):
    """Unit tests for equipment reference resolution and tokenization."""

    def setUp(self):
        self.resolver = IdentityResolver()

    def test_resolve_exact_huawei(self):
        candidates = self.resolver.resolve("Huawei SUN2000-100KTL-M1")
        self.assertTrue(len(candidates) >= 1)
        top = candidates[0]
        self.assertEqual(top.manufacturer, "Huawei")
        self.assertIn("SUN2000-100KTL-M1", top.exact_model)

    def test_resolve_alias_aiswei_solplanet(self):
        candidates = self.resolver.resolve("Aiswei Ai-Logger")
        self.assertTrue(len(candidates) >= 1)
        top = candidates[0]
        self.assertEqual(top.manufacturer, "Solplanet")
        self.assertEqual(top.exact_model, "Ai-Logger")

    def test_resolve_foxess_and_dyness(self):
        fox_candidates = self.resolver.resolve("Fox ESS H3-Pro 15.0")
        self.assertEqual(fox_candidates[0].manufacturer, "FoxESS")

        dyn_candidates = self.resolver.resolve("Dyness DH200Y")
        self.assertEqual(dyn_candidates[0].manufacturer, "Dyness")
        self.assertEqual(dyn_candidates[0].exact_model, "DH200Y")

    def test_resolve_empty_query(self):
        self.assertEqual(self.resolver.resolve(""), [])
        self.assertEqual(self.resolver.resolve("   "), [])
