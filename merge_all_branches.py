#!/usr/bin/env python3
"""
Script to merge all branches into main branch.
Handles merge conflicts and provides detailed reporting.
"""

import subprocess
import sys
from typing import List, Tuple

def run_command(cmd: List[str], cwd: str = ".") -> Tuple[int, str, str]:
    """Execute a shell command and return exit code, stdout, stderr."""
    try:
        result = subprocess.run(
            cmd,
            cwd=cwd,
            capture_output=True,
            text=True,
            timeout=30
        )
        return result.returncode, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return 1, "", "Command timed out"
    except Exception as e:
        return 1, "", str(e)

def get_all_branches() -> List[str]:
    """Get all branches except main."""
    returncode, stdout, stderr = run_command(["git", "branch", "-a"])
    if returncode != 0:
        print(f"Error getting branches: {stderr}")
        return []
    
    branches = []
    for line in stdout.split('\n'):
        line = line.strip()
        if not line or line.startswith('*'):
            continue
        # Remove remote tracking prefix
        if line.startswith('remotes/origin/'):
            branch = line.replace('remotes/origin/', '')
        else:
            branch = line
        
        if branch != 'main' and branch != 'HEAD':
            branches.append(branch)
    
    return sorted(set(branches))  # Remove duplicates and sort

def ensure_main_branch() -> bool:
    """Switch to main branch and ensure it's up to date."""
    print("Switching to main branch...")
    returncode, _, stderr = run_command(["git", "checkout", "main"])
    if returncode != 0:
        print(f"Error checking out main: {stderr}")
        return False
    
    print("Pulling latest changes to main...")
    returncode, _, stderr = run_command(["git", "pull", "origin", "main"])
    if returncode != 0:
        print(f"Warning: Could not pull main: {stderr}")
    
    return True

def merge_branch(branch: str) -> Tuple[bool, str]:
    """Attempt to merge a branch into main."""
    print(f"\nMerging {branch}...", end=" ")
    
    # Fetch the branch if it's remote
    if '/' not in branch:
        run_command(["git", "fetch", "origin", branch])
    
    returncode, stdout, stderr = run_command(["git", "merge", "--no-edit", branch])
    
    if returncode == 0:
        print("✓ Success")
        return True, "Merged successfully"
    else:
        # Check if it's a conflict
        if "CONFLICT" in stdout or "CONFLICT" in stderr:
            print("✗ Conflict")
            # Abort the merge
            run_command(["git", "merge", "--abort"])
            return False, "Merge conflict (aborted)"
        else:
            print("✗ Failed")
            # Try to abort
            run_command(["git", "merge", "--abort"])
            return False, f"Merge failed: {stderr[:100]}"

def main():
    """Main merge orchestration."""
    print("=== Branch Merge Automation ===\n")
    
    # Ensure we're in the repo root
    returncode, _, _ = run_command(["git", "rev-parse", "--git-dir"])
    if returncode != 0:
        print("Error: Not in a git repository")
        sys.exit(1)
    
    # Switch to main
    if not ensure_main_branch():
        print("Error: Could not switch to main branch")
        sys.exit(1)
    
    # Get all branches
    branches = get_all_branches()
    if not branches:
        print("No branches to merge")
        sys.exit(0)
    
    print(f"\nFound {len(branches)} branches to merge\n")
    
    # Track results
    successful = []
    failed = []
    
    # Merge each branch
    for i, branch in enumerate(branches, 1):
        print(f"[{i}/{len(branches)}]", end=" ")
        success, message = merge_branch(branch)
        
        if success:
            successful.append(branch)
        else:
            failed.append((branch, message))
    
    # Summary
    print("\n" + "="*50)
    print("MERGE SUMMARY")
    print("="*50)
    print(f"Total branches: {len(branches)}")
    print(f"Successful: {len(successful)}")
    print(f"Failed: {len(failed)}")
    
    if successful:
        print(f"\n✓ Merged ({len(successful)}):")
        for branch in successful[:10]:
            print(f"  - {branch}")
        if len(successful) > 10:
            print(f"  ... and {len(successful) - 10} more")
    
    if failed:
        print(f"\n✗ Failed ({len(failed)}):")
        for branch, reason in failed[:10]:
            print(f"  - {branch}: {reason}")
        if len(failed) > 10:
            print(f"  ... and {len(failed) - 10} more")
    
    # Push changes
    if successful:
        print("\nPushing merged changes to origin main...")
        returncode, stdout, stderr = run_command(["git", "push", "origin", "main"])
        if returncode == 0:
            print("✓ Push successful")
        else:
            print(f"✗ Push failed: {stderr}")
    
    print("\n" + "="*50)
    
    # Exit with appropriate code
    sys.exit(0 if not failed else 1)

if __name__ == "__main__":
    main()
