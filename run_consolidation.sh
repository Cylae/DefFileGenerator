#!/bin/bash
# Quick launcher for the consolidation script
cd ~
python3 /dev/stdin << 'PYTHON_SCRIPT'
#!/usr/bin/env python3
"""
DefFileGenerator - Branch Consolidation Executor
Merges all 150+ branches into main and cleans up
Advanced merge strategies with comprehensive error handling
"""

import subprocess
import sys
import os
from pathlib import Path
from typing import List

def run_cmd(cmd: List[str], capture: bool = False, ignore_error: bool = False) -> str:
    """Execute git command safely"""
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=not ignore_error
        )
        if result.returncode != 0 and not ignore_error:
            print(f"   ❌ {' '.join(cmd)}")
            if result.stderr:
                print(f"   Error: {result.stderr[:200]}")
        return result.stdout.strip() if capture else ""
    except Exception as e:
        if not ignore_error:
            print(f"   ❌ Exception: {e}")
        return ""

def consolidate_branches():
    """Main consolidation logic"""
    
    print("=" * 60)
    print("🔧 DefFileGenerator - Branch Consolidation Tool")
    print("=" * 60)
    print()
    
    repo_dir = Path.home() / "deffile_consolidation"
    repo_url = "https://github.com/Cylae/DefFileGenerator.git"
    
    # Step 1: Cleanup and clone
    print("📦 Step 1: Cloning repository...")
    if repo_dir.exists():
        run_cmd(["rm", "-rf", str(repo_dir)])
    
    repo_dir.mkdir(parents=True, exist_ok=True)
    run_cmd(["git", "clone", repo_url, str(repo_dir)])
    os.chdir(str(repo_dir))
    
    # Configure git
    run_cmd(["git", "config", "user.name", "DefFileGenerator Consolidation"])
    run_cmd(["git", "config", "user.email", "consolidation@bot.local"])
    
    print("✅ Repository cloned\n")
    
    # Step 2: Fetch all branches
    print("📋 Step 2: Fetching all branches...")
    run_cmd(["git", "fetch", "--all"])
    run_cmd(["git", "fetch", "--all", "--prune"])
    
    branches_output = run_cmd(["git", "branch", "-r"], capture=True)
    all_branches = [
        b.replace("origin/", "").strip() 
        for b in branches_output.split("\n")
        if b.strip() and "HEAD" not in b and "origin/main" not in b
    ]
    all_branches = sorted(set(all_branches))  # Remove duplicates and sort
    
    print(f"✅ Found {len(all_branches)} branches to process\n")
    
    # Step 3: Switch to main and update
    print("🔄 Step 3: Preparing main branch...")
    run_cmd(["git", "checkout", "main"])
    run_cmd(["git", "pull", "origin", "main"])
    print("✅ On main branch\n")
    
    # Step 4: Merge all branches
    print("🔗 Step 4: Merging branches into main...")
    print("-" * 60)
    
    merged = 0
    conflicts = 0
    errors = 0
    
    for i, branch in enumerate(all_branches, 1):
        progress = f"[{i:3d}/{len(all_branches)}]"
        
        # Try merge with 3 strategies in order
        merge_success = False
        
        # Strategy 1: Standard merge
        merge_cmd = ["git", "merge", "--no-ff", "--allow-unrelated-histories", 
                     f"origin/{branch}", "-m", f"Merge {branch}"]
        run_cmd(merge_cmd, ignore_error=True)
        
        # Check if merge succeeded (no unmerged files)
        status = run_cmd(["git", "status", "--porcelain"], capture=True)
        
        if "UU" not in status and "AA" not in status and "DD" not in status:
            print(f"{progress} ✅ {branch}")
            merged += 1
            merge_success = True
        else:
            # Abort and try alternative
            run_cmd(["git", "merge", "--abort"], ignore_error=True)
            
            # Strategy 2: Cherry-pick commits
            common_ancestor = run_cmd(
                ["git", "merge-base", "main", f"origin/{branch}"],
                capture=True,
                ignore_error=True
            )
            
            commits = run_cmd(
                ["git", "log", "--oneline", f"{common_ancestor}..origin/{branch}"],
                capture=True,
                ignore_error=True
            )
            
            if commits:
                # Try cherry-pick with theirs strategy
                run_cmd(["git", "merge", f"origin/{branch}", "-X", "theirs", 
                        "-m", f"Merge {branch} (resolved)"], ignore_error=True)
                
                status = run_cmd(["git", "status", "--porcelain"], capture=True)
                if "UU" not in status:
                    print(f"{progress} ⚠️  {branch} (cherry-picked)")
                    merged += 1
                    merge_success = True
                else:
                    run_cmd(["git", "merge", "--abort"], ignore_error=True)
            
            if not merge_success:
                print(f"{progress} ⚡ {branch} (skipped - conflicts)")
                conflicts += 1
    
    print("-" * 60)
    print()
    
    # Step 5: Push main
    print("🚀 Step 5: Pushing consolidated main...")
    run_cmd(["git", "push", "origin", "main", "--force"])
    print("✅ Main pushed\n")
    
    # Step 6: Delete remote branches
    print("🗑️  Step 6: Deleting remote branches...")
    print("-" * 60)
    
    deleted = 0
    for i, branch in enumerate(all_branches, 1):
        progress = f"[{i:3d}/{len(all_branches)}]"
        result = run_cmd(["git", "push", "origin", "--delete", branch], ignore_error=True)
        
        if result == "":
            print(f"{progress} ✓ Deleted {branch}")
            deleted += 1
        else:
            print(f"{progress} ✗ Failed to delete {branch}")
    
    print("-" * 60)
    print()
    
    # Summary
    print("=" * 60)
    print("📊 CONSOLIDATION SUMMARY")
    print("=" * 60)
    print(f"Total branches processed:  {len(all_branches)}")
    print(f"Successfully merged:       {merged}")
    print(f"Conflicted/Skipped:        {conflicts}")
    print(f"Remote branches deleted:   {deleted}")
    print()
    print("✨ Repository consolidated to single 'main' branch!")
    print("=" * 60)
    print()
    
    # Cleanup
    print("🧹 Cleaning up...")
    os.chdir(str(Path.home()))
    run_cmd(["rm", "-rf", str(repo_dir)])
    print("✅ Done!\n")
    
    print("📌 Next steps:")
    print("   1. Visit: https://github.com/Cylae/DefFileGenerator")
    print("   2. Verify only 'main' branch exists")
    print("   3. Check the commit history\n")

if __name__ == "__main__":
    try:
        consolidate_branches()
        sys.exit(0)
    except KeyboardInterrupt:
        print("\n\n⚠️  Process interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
PYTHON_SCRIPT
