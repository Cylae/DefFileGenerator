"""
Security-Guarded Document Fetcher for Autonomous Modbus Discovery.
Enforces strict anti-SSRF protections, DNS pinning, IP validation,
redirect boundaries, streaming byte bounds, and SHA-256 verification.
"""

from __future__ import annotations

import hashlib
import ipaddress
import logging
import socket
import urllib.parse

import httpx

from DefFileGenerator.discovery.models import DocumentCandidate, FetchOutcome, FetchResult

logger = logging.getLogger(__name__)

MAX_DOWNLOAD_BYTES = 25 * 1024 * 1024  # 25 MB maximum
MAX_REDIRECTS = 5
DEFAULT_TIMEOUT_SEC = 30.0
ALLOWED_SCHEMES = frozenset({"http", "https"})


def is_ip_allowed(ip_str: str) -> bool:
    """
    Validates whether an IP address is a safe public routable address.
    Rejects private networks, loopback, link-local, multicast, and cloud metadata IPs.
    """
    try:
        ip = ipaddress.ip_address(ip_str)
        if ip.is_loopback:
            return False
        if ip.is_private:
            return False
        if ip.is_link_local:
            return False
        if ip.is_multicast:
            return False
        if ip.is_reserved:
            return False
        if ip.is_unspecified:
            return False
        # Specific check for AWS / GCP / Azure metadata service
        if str(ip) == "169.254.169.254":
            return False
        return True
    except ValueError:
        return False


def validate_target_url(url: str) -> tuple[bool, str, list[str]]:
    """
    Validates a URL against SSRF and network access policies.
    Returns (is_valid, error_reason, resolved_ips).
    """
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme.lower() not in ALLOWED_SCHEMES:
        return (
            False,
            f"Forbidden URL scheme '{parsed.scheme}'. Only HTTPS and HTTP are allowed.",
            [],
        )

    hostname = parsed.hostname
    if not hostname:
        return False, "URL contains no valid hostname.", []

    # Check for localhost / loopback aliases
    if hostname.lower() in ("localhost", "127.0.0.1", "::1", "metadata.google.internal"):
        return False, f"Forbidden host '{hostname}' (loopback / internal).", []

    resolved_ips: list[str] = []
    try:
        addr_info = socket.getaddrinfo(hostname, None)
        for item in addr_info:
            ip = str(item[4][0])
            if ip not in resolved_ips:
                resolved_ips.append(ip)
    except Exception as e:
        return False, f"DNS resolution failed for '{hostname}': {e}", []

    if not resolved_ips:
        return False, f"Could not resolve any IP address for host '{hostname}'.", []

    for ip_str in resolved_ips:
        if not is_ip_allowed(ip_str):
            return (
                False,
                f"Forbidden destination IP '{ip_str}' resolved from '{hostname}'.",
                resolved_ips,
            )

    return True, "", resolved_ips


class DocumentFetcher:
    """SSRF-safe HTTP client with streaming limits and SHA-256 verification."""

    def __init__(
        self,
        max_bytes: int = MAX_DOWNLOAD_BYTES,
        max_redirects: int = MAX_REDIRECTS,
        timeout_sec: float = DEFAULT_TIMEOUT_SEC,
    ) -> None:
        self.max_bytes = max_bytes
        self.max_redirects = max_redirects
        self.timeout_sec = timeout_sec

    async def fetch(
        self,
        candidate: DocumentCandidate | str,
        custom_headers: dict[str, str] | None = None,
    ) -> FetchResult:
        """
        Retrieves a document safely from a URL or DocumentCandidate.
        Enforces SSRF boundary guards at every redirect hop.
        """
        current_url = (
            candidate.source_url if isinstance(candidate, DocumentCandidate) else candidate
        )
        headers = {
            "User-Agent": "WebdynModbusDiscovery/1.0 (Industrial Protocol Ingestion Engine; +https://webdyn.com)",
            "Accept": "application/pdf, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, text/csv, text/html, */*",
        }
        if custom_headers:
            headers.update(custom_headers)

        redirects_followed = 0
        while redirects_followed <= self.max_redirects:
            # 1. Pre-flight SSRF check on target URL
            is_valid, reason, _ = validate_target_url(current_url)
            if not is_valid:
                logger.warning("SSRF block on target URL %s: %s", current_url, reason)
                return FetchResult(
                    outcome=FetchOutcome.BLOCKED_SSRF,
                    final_url=current_url,
                    error_message=f"SSRF guard violation: {reason}",
                )

            # 2. Issue streaming request with follow_redirects=False for manual validation
            try:
                async with httpx.AsyncClient(
                    verify=True,
                    timeout=httpx.Timeout(self.timeout_sec, connect=10.0),
                    follow_redirects=False,
                ) as client:
                    async with client.stream("GET", current_url, headers=headers) as resp:
                        # Check redirects
                        if resp.is_redirect:
                            location = resp.headers.get("location")
                            if not location:
                                return FetchResult(
                                    outcome=FetchOutcome.ACCESS_RESTRICTED,
                                    final_url=current_url,
                                    error_message="Redirect received without Location header.",
                                )
                            current_url = urllib.parse.urljoin(current_url, location)
                            redirects_followed += 1
                            continue

                        if resp.status_code == 429:
                            return FetchResult(
                                outcome=FetchOutcome.RATE_LIMITED,
                                final_url=current_url,
                                error_message="Source server returned HTTP 429 Rate Limited.",
                            )
                        if resp.status_code in (401, 403):
                            return FetchResult(
                                outcome=FetchOutcome.ACCESS_RESTRICTED,
                                final_url=current_url,
                                error_message=f"HTTP {resp.status_code} Access Restricted.",
                            )
                        if resp.status_code != 200:
                            return FetchResult(
                                outcome=FetchOutcome.ACCESS_RESTRICTED,
                                final_url=current_url,
                                error_message=f"HTTP {resp.status_code} returned by host.",
                            )

                        # Content-Length header guard
                        content_length = resp.headers.get("content-length")
                        if content_length and content_length.isdigit():
                            if int(content_length) > self.max_bytes:
                                return FetchResult(
                                    outcome=FetchOutcome.OVERSPEC_BYTES,
                                    final_url=current_url,
                                    error_message=f"Content-Length {content_length} exceeds max allowed {self.max_bytes} bytes.",
                                )

                        # Stream content with byte accumulator and live hash computation
                        chunks: list[bytes] = []
                        total_bytes = 0
                        hasher = hashlib.sha256()

                        async for chunk in resp.aiter_bytes(chunk_size=65536):
                            total_bytes += len(chunk)
                            if total_bytes > self.max_bytes:
                                return FetchResult(
                                    outcome=FetchOutcome.OVERSPEC_BYTES,
                                    final_url=current_url,
                                    error_message=f"Download stream exceeded {self.max_bytes} bytes.",
                                )
                            chunks.append(chunk)
                            hasher.update(chunk)

                        full_bytes = b"".join(chunks)
                        content_type = resp.headers.get("content-type", "application/octet-stream")
                        sha256_hash = hasher.hexdigest()

                        return FetchResult(
                            outcome=FetchOutcome.SUCCESS,
                            file_bytes=full_bytes,
                            sha256=sha256_hash,
                            content_type=content_type,
                            final_url=current_url,
                        )

            except httpx.TimeoutException as e:
                return FetchResult(
                    outcome=FetchOutcome.NETWORK_TIMEOUT,
                    final_url=current_url,
                    error_message=f"Network timeout ({self.timeout_sec}s): {e}",
                )
            except Exception as e:
                return FetchResult(
                    outcome=FetchOutcome.ACCESS_RESTRICTED,
                    final_url=current_url,
                    error_message=f"Network error during retrieval: {e}",
                )

        return FetchResult(
            outcome=FetchOutcome.ACCESS_RESTRICTED,
            final_url=current_url,
            error_message=f"Exceeded maximum redirects ({self.max_redirects}).",
        )
