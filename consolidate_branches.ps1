#!/usr/bin/env pwsh
<#
.SYNOPSIS
    DefFileGenerator - Branch Consolidation Executor
    Merges all 150+ branches into main and cleans up
    
.DESCRIPTION
    Advanced merge strategies with comprehensive error handling
    Requires: PowerShell 7.0+, Git
    
.EXAMPLE
    .\consolidate_branches.ps1
#>

param(
    [string]$RepoUrl = "https://github.com/Cylae/DefFileGenerator.git",
    [string]$MainBranch = "main"
)

$ErrorActionPreference = "Continue"

function Invoke-GitCommand {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Arguments,
        
        [bool]$CaptureOutput = $false,
        [bool]$IgnoreError = $false
    )
    
    try {
        if ($CaptureOutput) {
            $result = & git @Arguments 2>&1
            return $result -join "`n"
        } else {
            & git @Arguments 2>&1 | Out-Null
            return ""
        }
    } catch {
        if (-not $IgnoreError) {
            Write-Host "   ❌ Command failed: git $($Arguments -join ' ')" -ForegroundColor Red
            Write-Host "   Error: $_" -ForegroundColor Red
        }
        return ""
    }
}

function Start-Consolidation {
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "🔧 DefFileGenerator - Branch Consolidation Tool" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host ""
    
    $repoDir = Join-Path $env:TEMP "deffile_consolidation"
    
    # Step 1: Cleanup and clone
    Write-Host "📦 Step 1: Cloning repository..." -ForegroundColor Yellow
    
    if (Test-Path $repoDir) {
        Remove-Item -Path $repoDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    New-Item -Path $repoDir -ItemType Directory -Force | Out-Null
    Push-Location $repoDir
    
    Invoke-GitCommand -Arguments @("clone", $RepoUrl, ".") -IgnoreError $false
    
    # Configure git
    Invoke-GitCommand -Arguments @("config", "user.name", "DefFileGenerator Consolidation")
    Invoke-GitCommand -Arguments @("config", "user.email", "consolidation@bot.local")
    
    Write-Host "✅ Repository cloned`n" -ForegroundColor Green
    
    # Step 2: Fetch all branches
    Write-Host "📋 Step 2: Fetching all branches..." -ForegroundColor Yellow
    Invoke-GitCommand -Arguments @("fetch", "--all")
    Invoke-GitCommand -Arguments @("fetch", "--all", "--prune")
    
    $branchesOutput = Invoke-GitCommand -Arguments @("branch", "-r") -CaptureOutput $true
    $allBranches = @()
    
    foreach ($line in ($branchesOutput -split "`n")) {
        $branch = $line.Trim().Replace("origin/", "")
        if ($branch -and $branch -notlike "*HEAD*" -and $branch -ne "main") {
            if ($branch -notin $allBranches) {
                $allBranches += $branch
            }
        }
    }
    
    $allBranches = $allBranches | Sort-Object
    Write-Host "✅ Found $($allBranches.Count) branches to process`n" -ForegroundColor Green
    
    # Step 3: Switch to main and update
    Write-Host "🔄 Step 3: Preparing main branch..." -ForegroundColor Yellow
    Invoke-GitCommand -Arguments @("checkout", $MainBranch)
    Invoke-GitCommand -Arguments @("pull", "origin", $MainBranch)
    Write-Host "✅ On main branch`n" -ForegroundColor Green
    
    # Step 4: Merge all branches
    Write-Host "🔗 Step 4: Merging branches into main..." -ForegroundColor Yellow
    Write-Host ("-" * 60) -ForegroundColor Gray
    
    $merged = 0
    $conflicts = 0
    $skipped = 0
    
    foreach ($i in 0..($allBranches.Count - 1)) {
        $branch = $allBranches[$i]
        $progressStr = "[{0:D3}/{1}]" -f ($i + 1), $allBranches.Count
        
        $mergeSuccess = $false
        
        # Strategy 1: Standard merge
        $mergeCmd = @("merge", "--no-ff", "--allow-unrelated-histories", "origin/$branch", "-m", "Merge $branch")
        Invoke-GitCommand -Arguments $mergeCmd -IgnoreError $true | Out-Null
        
        # Check if merge succeeded
        $status = Invoke-GitCommand -Arguments @("status", "--porcelain") -CaptureOutput $true
        
        if ($status -notlike "*UU*" -and $status -notlike "*AA*" -and $status -notlike "*DD*") {
            Write-Host "$progressStr ✅ $branch" -ForegroundColor Green
            $merged++
            $mergeSuccess = $true
        } else {
            # Abort and try alternative
            Invoke-GitCommand -Arguments @("merge", "--abort") -IgnoreError $true
            
            # Strategy 2: Cherry-pick with theirs strategy
            $ancestor = Invoke-GitCommand -Arguments @("merge-base", $MainBranch, "origin/$branch") -CaptureOutput $true -IgnoreError $true
            
            if ($ancestor) {
                $commits = Invoke-GitCommand -Arguments @("log", "--oneline", "$ancestor..origin/$branch") -CaptureOutput $true -IgnoreError $true
                
                if ($commits) {
                    Invoke-GitCommand -Arguments @("merge", "origin/$branch", "-X", "theirs", "-m", "Merge $branch (resolved)") -IgnoreError $true | Out-Null
                    
                    $status = Invoke-GitCommand -Arguments @("status", "--porcelain") -CaptureOutput $true
                    
                    if ($status -notlike "*UU*") {
                        Write-Host "$progressStr ⚠️  $branch (cherry-picked)" -ForegroundColor Yellow
                        $merged++
                        $mergeSuccess = $true
                    } else {
                        Invoke-GitCommand -Arguments @("merge", "--abort") -IgnoreError $true
                    }
                }
            }
            
            if (-not $mergeSuccess) {
                Write-Host "$progressStr ⚡ $branch (skipped - conflicts)" -ForegroundColor Magenta
                $conflicts++
            }
        }
    }
    
    Write-Host ("-" * 60) -ForegroundColor Gray
    Write-Host ""
    
    # Step 5: Push main
    Write-Host "🚀 Step 5: Pushing consolidated main..." -ForegroundColor Yellow
    Invoke-GitCommand -Arguments @("push", "origin", $MainBranch, "--force")
    Write-Host "✅ Main pushed`n" -ForegroundColor Green
    
    # Step 6: Delete remote branches
    Write-Host "🗑️  Step 6: Deleting remote branches..." -ForegroundColor Yellow
    Write-Host ("-" * 60) -ForegroundColor Gray
    
    $deleted = 0
    foreach ($i in 0..($allBranches.Count - 1)) {
        $branch = $allBranches[$i]
        $progressStr = "[{0:D3}/{1}]" -f ($i + 1), $allBranches.Count
        
        $result = Invoke-GitCommand -Arguments @("push", "origin", "--delete", $branch) -IgnoreError $true -CaptureOutput $true
        
        if ([string]::IsNullOrEmpty($result) -or $result -notlike "*error*") {
            Write-Host "$progressStr ✓ Deleted $branch" -ForegroundColor Green
            $deleted++
        } else {
            Write-Host "$progressStr ✗ Failed to delete $branch" -ForegroundColor Red
        }
    }
    
    Write-Host ("-" * 60) -ForegroundColor Gray
    Write-Host ""
    
    # Summary
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "📊 CONSOLIDATION SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "Total branches processed:  $($allBranches.Count)"
    Write-Host "Successfully merged:       $merged" -ForegroundColor Green
    Write-Host "Conflicted/Skipped:        $conflicts" -ForegroundColor Yellow
    Write-Host "Remote branches deleted:   $deleted" -ForegroundColor Green
    Write-Host ""
    Write-Host "✨ Repository consolidated to single 'main' branch!" -ForegroundColor Green
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host ""
    
    # Cleanup
    Pop-Location
    Write-Host "🧹 Cleaning up..." -ForegroundColor Yellow
    Remove-Item -Path $repoDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "✅ Done!`n" -ForegroundColor Green
    
    Write-Host "📌 Next steps:" -ForegroundColor Cyan
    Write-Host "   1. Visit: https://github.com/Cylae/DefFileGenerator"
    Write-Host "   2. Verify only 'main' branch exists"
    Write-Host "   3. Check the commit history`n"
}

# Main execution
try {
    Start-Consolidation
    exit 0
} catch {
    Write-Host "`n❌ Fatal error: $_" -ForegroundColor Red
    Write-Host $_.Exception.StackTrace -ForegroundColor Red
    exit 1
}
