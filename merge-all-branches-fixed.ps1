#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Fixed GitHub Branch Merge Script with improved error handling and conflict resolution
.DESCRIPTION
    Merges all branches into main using GitHub CLI with:
    - Squash merging for conflict handling
    - Proper error handling and logging
    - Log file in user directory (not system directory)
    - Better PR creation and merge strategies
#>

param(
    [string]$Owner = "Cylae",
    [string]$Repo = "DefFileGenerator",
    [string]$TargetBranch = "main",
    [switch]$UseSquashMerge = $true,
    [switch]$Force = $false
)

# Get user's home directory for log file
$LogDir = Join-Path $env:USERPROFILE "Desktop"
if (-not (Test-Path $LogDir)) {
    $LogDir = $env:USERPROFILE
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $LogDir "merge_log_${Timestamp}.txt"

# Initialize log file
function Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Write-Host $LogMessage
    Add-Content -Path $LogFile -Value $LogMessage -ErrorAction SilentlyContinue
}

# Header
Log "====================================================================="
Log "GitHub Branch Merge Script (FIXED)"
Log "Repository: $Owner/$Repo"
Log "Target Branch: $TargetBranch"
Log "Merge Strategy: $(if ($UseSquashMerge) { 'Squash' } else { 'Regular' })"
Log "Log file: $LogFile"
Log "====================================================================="

# Check GitHub CLI
Log "Checking for GitHub CLI installation..."
$ghCheck = gh --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: GitHub CLI not found" "ERROR"
    exit 1
}
Log "✓ GitHub CLI found: $ghCheck"

# Check authentication
Log "Checking GitHub authentication..."
$authCheck = gh auth status 2>&1
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: Not authenticated with GitHub" "ERROR"
    Log "Run: gh auth login"
    exit 1
}
Log "✓ Authenticated with GitHub"

# Verify repository exists
Log "Verifying repository access..."
$repoCheck = gh repo view "$Owner/$Repo" --json nameWithOwner 2>&1
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: Cannot access repository $Owner/$Repo" "ERROR"
    exit 1
}
Log "✓ Repository accessible"

# Create temp directory for repo
$tempRepo = Join-Path $env:TEMP "DefFileGenerator_merge_$$"

# Fetch all branches
Log "Fetching all branches from remote..."
gh repo clone "$Owner/$Repo" $tempRepo --depth=1 -q 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: Could not clone repository" "ERROR"
    exit 1
}

& git -C $tempRepo fetch --all -q 2>&1 | Out-Null
Log "✓ Branches fetched"

# Get all branches except main
Log "Collecting branches to merge..."
$allBranches = @()
$remoteBranches = & git -C $tempRepo branch -r 2>&1

foreach ($branch in $remoteBranches) {
    $trimmed = $branch.Trim()
    # Skip HEAD and main
    if ($trimmed -match "HEAD|main|master" -or $trimmed -eq "") {
        continue
    }
    # Remove origin/ prefix
    $cleanBranch = $trimmed -replace "^origin/", ""
    $allBranches += $cleanBranch
}

$uniqueBranches = $allBranches | Sort-Object -Unique
Log "✓ Total branches found: $($uniqueBranches.Count)"

if ($uniqueBranches.Count -eq 0) {
    Log "No branches to merge"
    Remove-Item -Path $tempRepo -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    exit 0
}

# Merge tracking
$merged = @()
$failed = @()
$skipped = @()

Log ""
Log "Starting merge operations..."
Log ""

$counter = 0
foreach ($branch in $uniqueBranches) {
    $counter++
    Log "[${counter}/$($uniqueBranches.Count)] Processing branch: $branch"
    
    # Skip if branch doesn't exist
    $branchExists = & git -C $tempRepo ls-remote --heads origin "$branch" 2>&1
    if (-not $branchExists) {
        Log "  ✗ Branch not found on remote"
        $skipped += $branch
        continue
    }
    
    # Create PR
    Log "  → Creating PR..."
    $prOutput = gh pr create --repo "$Owner/$Repo" --head "$branch" --base "$TargetBranch" --title "Merge $branch into $TargetBranch" --body "Automated merge" --fill 2>&1
    
    $prNumber = $null
    
    if ($LASTEXITCODE -ne 0) {
        # PR might already exist, try to find it
        $existingPR = gh pr list --repo "$Owner/$Repo" --head "$branch" --base "$TargetBranch" --state open --json number 2>&1 | ConvertFrom-Json | Select-Object -First 1
        
        if ($existingPR -and $existingPR.number) {
            $prNumber = $existingPR.number
            Log "  ✓ Using existing PR #$prNumber"
        } else {
            Log "  ✗ Could not create or find PR"
            $skipped += $branch
            continue
        }
    } else {
        # Extract PR number from output
        if ($prOutput -match "#(\d+)") {
            $prNumber = [int]$matches[1]
            Log "  ✓ PR #$prNumber created"
        } else {
            Log "  ✗ Could not retrieve PR number"
            $skipped += $branch
            continue
        }
    }
    
    # Merge PR with retry logic
    Log "  → Merging PR..."
    $mergeAttempts = 0
    $maxAttempts = 3
    $mergeSuccess = $false
    
    while ($mergeAttempts -lt $maxAttempts -and -not $mergeSuccess) {
        $mergeAttempts++
        
        # Try merge with strategy
        if ($UseSquashMerge) {
            $mergeResult = gh pr merge $prNumber --repo "$Owner/$Repo" --squash --auto 2>&1
        } else {
            $mergeResult = gh pr merge $prNumber --repo "$Owner/$Repo" --merge --auto 2>&1
        }
        
        if ($LASTEXITCODE -eq 0) {
            $mergeSuccess = $true
            Log "  ✓ Successfully merged and branch deleted"
            $merged += $branch
            break
        } else {
            $errorMsg = $mergeResult -join " "
            
            # Check for specific error types
            if ($errorMsg -match "merge conflict|CONFLICT") {
                if ($mergeAttempts -lt $maxAttempts) {
                    Log "  ! Merge conflict detected, retrying with squash..."
                    # Force squash on retry
                    $retryMerge = gh pr merge $prNumber --repo "$Owner/$Repo" --squash --force 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        $mergeSuccess = $true
                        Log "  ✓ Successfully merged with squash"
                        $merged += $branch
                        break
                    }
                } else {
                    Log "  ✗ Merge conflicts remain (requires manual resolution)"
                    $failed += @{branch = $branch; reason = "Merge conflicts"; prNumber = $prNumber}
                    break
                }
            } elseif ($errorMsg -match "Status: DIRTY|not ready|cannot be merged") {
                Log "  ! PR status not ready, waiting..."
                Start-Sleep -Seconds 2
            } else {
                $shortError = if ($errorMsg.Length -gt 80) { $errorMsg.Substring(0, 80) + "..." } else { $errorMsg }
                Log "  ✗ Merge failed: $shortError"
                $failed += @{branch = $branch; reason = $shortError; prNumber = $prNumber}
                break
            }
        }
    }
}

# Cleanup temp repo
Log ""
Log "Cleaning up temporary repository..."
Remove-Item -Path $tempRepo -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
Log "✓ Cleanup complete"

# Summary
Log ""
Log "====================================================================="
Log "MERGE SUMMARY"
Log "====================================================================="
Log "Total branches processed: $($uniqueBranches.Count)"
Log "Successfully merged: $($merged.Count)"
Log "Failed: $($failed.Count)"
Log "Skipped: $($skipped.Count)"

if ($merged.Count -gt 0) {
    Log ""
    Log "✓ MERGED ($($merged.Count)):"
    foreach ($branch in $merged | Select-Object -First 10) {
        Log "  - $branch"
    }
    if ($merged.Count -gt 10) {
        Log "  ... and $($merged.Count - 10) more"
    }
}

if ($failed.Count -gt 0) {
    Log ""
    Log "✗ FAILED ($($failed.Count)):"
    foreach ($item in $failed | Select-Object -First 10) {
        $reason = if ($item.reason.Length -gt 50) { $item.reason.Substring(0, 50) + "..." } else { $item.reason }
        Log "  - $($item.branch) (PR #$($item.prNumber)): $reason"
    }
    if ($failed.Count -gt 10) {
        Log "  ... and $($failed.Count - 10) more"
    }
}

if ($skipped.Count -gt 0) {
    Log ""
    Log "⊘ SKIPPED ($($skipped.Count)):"
    foreach ($branch in $skipped | Select-Object -First 10) {
        Log "  - $branch"
    }
    if ($skipped.Count -gt 10) {
        Log "  ... and $($skipped.Count - 10) more"
    }
}

Log ""
Log "====================================================================="
Log "Process completed!"
Log "View detailed log: $LogFile"
Log "====================================================================="

if ($failed.Count -gt 0) {
    Log ""
    Log "NEXT STEPS for failed merges:"
    Log "1. Review conflicts at: https://github.com/$Owner/$Repo/pulls"
    Log "2. Resolve conflicts manually in the GitHub UI"
    Log "3. Re-run this script to retry remaining branches"
    exit 1
} else {
    exit 0
}
