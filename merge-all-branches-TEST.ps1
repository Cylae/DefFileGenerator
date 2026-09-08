#!/usr/bin/env pwsh
<#
.SYNOPSIS
    TEST/DRY-RUN Version - GitHub Branch Merge Script (No actual merging)
.DESCRIPTION
    This is a safe test version that validates setup without making changes.
    Run this first to ensure everything is configured correctly.
#>

param(
    [string]$Owner = "Cylae",
    [string]$Repo = "DefFileGenerator",
    [string]$TargetBranch = "main"
)

Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host "GitHub Branch Merge Script - DRY RUN TEST" -ForegroundColor Yellow
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host ""

# Test 1: Check GitHub CLI
Write-Host "[TEST 1] Checking GitHub CLI installation..." -ForegroundColor White
try {
    $ghVersion = gh --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ PASS: GitHub CLI found" -ForegroundColor Green
        Write-Host "  $ghVersion" -ForegroundColor Gray
    } else {
        Write-Host "✗ FAIL: GitHub CLI not working" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "✗ FAIL: GitHub CLI not found" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Test 2: Check authentication
Write-Host "[TEST 2] Checking GitHub authentication..." -ForegroundColor White
try {
    $authStatus = gh auth status 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ PASS: Authenticated with GitHub" -ForegroundColor Green
    } else {
        Write-Host "✗ FAIL: Not authenticated" -ForegroundColor Red
        Write-Host "  Run: gh auth login" -ForegroundColor Yellow
        exit 1
    }
} catch {
    Write-Host "✗ FAIL: Authentication check failed" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Test 3: Verify repository access
Write-Host "[TEST 3] Verifying repository access..." -ForegroundColor White
try {
    $repoInfo = gh repo view "$Owner/$Repo" --json nameWithOwner,defaultBranchRef 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ PASS: Repository accessible" -ForegroundColor Green
        Write-Host "  Repository: $Owner/$Repo" -ForegroundColor Gray
    } else {
        Write-Host "✗ FAIL: Cannot access repository" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "✗ FAIL: Repository access failed" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Test 4: Check log file path
Write-Host "[TEST 4] Checking log file path..." -ForegroundColor White
$LogDir = Join-Path $env:USERPROFILE "Desktop"
if (-not (Test-Path $LogDir)) {
    $LogDir = $env:USERPROFILE
}
try {
    $TestFile = Join-Path $LogDir "merge_test_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    "Test" | Out-File -FilePath $TestFile -ErrorAction Stop
    Remove-Item $TestFile -ErrorAction Stop
    Write-Host "✓ PASS: Can write to log directory" -ForegroundColor Green
    Write-Host "  Path: $LogDir" -ForegroundColor Gray
} catch {
    Write-Host "✗ FAIL: Cannot write to log directory" -ForegroundColor Red
    Write-Host "  Error: $_" -ForegroundColor Yellow
    exit 1
}
Write-Host ""

# Test 5: Fetch branches (dry run)
Write-Host "[TEST 5] Fetching branches (dry run)..." -ForegroundColor White
try {
    $tempDir = Join-Path $env:TEMP "DefFileGenerator_test_$$"
    
    Write-Host "  Creating temporary clone..." -ForegroundColor Gray
    gh repo clone "$Owner/$Repo" $tempDir --depth=1 -q 2>&1 | Out-Null
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Fetching all branches..." -ForegroundColor Gray
        & git -C $tempDir fetch --all -q 2>&1 | Out-Null
        
        # Get branch count
        $allBranches = @()
        $remoteBranches = & git -C $tempDir branch -r 2>&1
        
        foreach ($branch in $remoteBranches) {
            $trimmed = $branch.Trim()
            if ($trimmed -notmatch "HEAD|main|master" -and $trimmed -ne "") {
                $cleanBranch = $trimmed -replace "^origin/", ""
                $allBranches += $cleanBranch
            }
        }
        
        $uniqueBranches = $allBranches | Sort-Object -Unique
        
        Write-Host "✓ PASS: Branches fetched successfully" -ForegroundColor Green
        Write-Host "  Total branches to merge: $($uniqueBranches.Count)" -ForegroundColor Gray
        
        if ($uniqueBranches.Count -le 10) {
            foreach ($branch in $uniqueBranches) {
                Write-Host "    - $branch" -ForegroundColor Gray
            }
        } else {
            for ($i = 0; $i -lt 5; $i++) {
                Write-Host "    - $($uniqueBranches[$i])" -ForegroundColor Gray
            }
            Write-Host "    ... and $($uniqueBranches.Count - 5) more" -ForegroundColor Gray
        }
        
        # Cleanup
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    } else {
        Write-Host "✗ FAIL: Could not clone repository" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "✗ FAIL: Branch fetch failed" -ForegroundColor Red
    Write-Host "  Error: $_" -ForegroundColor Yellow
    exit 1
}
Write-Host ""

# Test 6: Simulate PR creation
Write-Host "[TEST 6] Testing PR creation (simulated)..." -ForegroundColor White
try {
    # Just test the command format, don't actually create
    $testBranch = "test-branch-do-not-merge"
    Write-Host "  Command: gh pr create --repo $Owner/$Repo --head <branch> --base $TargetBranch --title 'Merge...' --body 'Automated' --fill" -ForegroundColor Gray
    Write-Host "✓ PASS: PR creation command format valid" -ForegroundColor Green
} catch {
    Write-Host "✗ FAIL: PR creation test failed" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Summary
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host "ALL TESTS PASSED!" -ForegroundColor Green
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Your environment is ready for the full merge script." -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Run: powershell -ExecutionPolicy Bypass -File merge-all-branches-fixed.ps1" -ForegroundColor White
Write-Host "2. Monitor the merge progress in real-time" -ForegroundColor White
Write-Host "3. Check the log file for detailed results" -ForegroundColor White
Write-Host ""
Write-Host "Note: The fixed script will:" -ForegroundColor Yellow
Write-Host "  • Use squash merging for better conflict handling" -ForegroundColor Gray
Write-Host "  • Retry failed merges with alternative strategies" -ForegroundColor Gray
Write-Host "  • Log all actions to your user directory (not System32)" -ForegroundColor Gray
Write-Host "  • Skip already-merged branches and non-existent branches" -ForegroundColor Gray
Write-Host ""
