<#
.SYNOPSIS
    Apollo-Standard Engineering Kit - PowerShell Bootstrap Script
.DESCRIPTION
    Creates the required `docs/audit/` durable ledgers and verifies repository setup.
#>

$ErrorActionPreference = "Stop"

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host " Apollo-Standard Engineering Kit - Initialization (PowerShell)" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

$auditDir = "docs/audit"
if (-not (Test-Path $auditDir)) {
    New-Item -ItemType Directory -Force -Path $auditDir | Out-Null
    Write-Host "[+] Created directory: $auditDir" -ForegroundColor Green
} else {
    Write-Host "[*] Directory already exists: $auditDir" -ForegroundColor Yellow
}

$ledgers = @("PLAN.md", "REVIEW_MATRIX.md", "FINDINGS.md", "VALIDATION.md", "PERFORMANCE.md", "HANDOFF.md")
foreach ($ledger in $ledgers) {
    $path = Join-Path $auditDir $ledger
    if (-not (Test-Path $path)) {
        # Create empty placeholder if not present in zip extraction
        Set-Content -Path $path -Value "# Placeholder for $ledger"
        Write-Host "[+] Initialized ledger: $path" -ForegroundColor Green
    } else {
        Write-Host "[*] Ledger exists: $path" -ForegroundColor Yellow
    }
}

Write-Host "`n[SUCCESS] Apollo-Standard environment initialized successfully!" -ForegroundColor Cyan
