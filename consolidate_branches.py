#!/usr/bin/env python3
"""
Branch consolidation script for DefFileGenerator
Merges all branches into main and deletes them
"""

import subprocess
import sys
import os
from typing import List, Tuple

class BranchConsolidator:
    def __init__(self, repo_owner: str, repo_name: str):
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.repo_url = f"https://github.com/{repo_owner}/{repo_name}.git"
        self.main_branch = "main"
        self.work_dir = f"{repo_name}-consolidation"
        self.merged_count = 0
        self.conflict_count = 0
        self.skipped_count = 0
        
    def run_cmd(self, cmd: List[str], check=True, capture=False) -> str:
        """Execute a shell command"""
        try:
            if capture:
                result = subprocess.run(cmd, check=check, capture_output=True, text=True)
                return result.stdout.strip()
            else:
                subprocess.run(cmd, check=check)
                return ""
        except subprocess.CalledProcessError as e:
            if check:
                print(f"❌ Command failed: {' '.join(cmd)}")
                print(f"   Error: {e.stderr if hasattr(e, 'stderr') else e}")
                raise
            return ""
    
    def setup_repo(self):
        """Clone and configure the repository"""
        print("🔄 Setting up repository...")
        
        # Clean up if exists
        if os.path.exists(self.work_dir):
            self.run_cmd(["rm", "-rf", self.work_dir])
        
        # Clone
        self.run_cmd(["git", "clone", self.repo_url, self.work_dir])
        os.chdir(self.work_dir)
        
        # Configure git
        self.run_cmd(["git", "config", "user.name", "Consolidation Bot"])
        self.run_cmd(["git", "config", "user.email", "bot@consolidation.local"])
        
        # Fetch all branches
        print("📋 Fetching all branches...")
        self.run_cmd(["git", "fetch", "--all"])
        
    def get_all_branches(self) -> List[str]:
        """Get all branches except main"""
        output = self.run_cmd(
            ["git", "branch", "-r"],
            capture=True
        )
        
        branches = []
        for line in output.split('\n'):
            line = line.strip()
            if not line or 'HEAD' in line or line == 'main':
                continue
            # Remove 'origin/' prefix
            branch = line.replace('origin/', '')
            if branch != 'main' and branch not in branches:
                branches.append(branch)
        
        return sorted(branches)
    
    def merge_branch(self, branch: str) -> bool:
        """Try to merge a branch into main"""
        print(f"\n⏳ Processing: {branch}")
        
        # Try standard merge
        try:
            self.run_cmd([
                "git", "merge", 
                "--no-ff", 
                "--allow-unrelated-histories",
                f"origin/{branch}",
                "-m", f"Merge branch '{branch}' into main"
            ], check=False)
            
            # Check if merge was successful
            status = self.run_cmd(["git", "status", "--porcelain"], capture=True)
            if not status or "MERGE" not in status:
                print(f"   ✅ Merged successfully")
                return True
            else:
                # Merge conflicts - abort
                print(f"   ⚠️  Conflicts detected, aborting...")
                self.run_cmd(["git", "merge", "--abort"], check=False)
                return False
                
        except Exception as e:
            print(f"   ✗ Merge failed: {e}")
            self.run_cmd(["git", "merge", "--abort"], check=False)
            return False
    
    def consolidate(self):
        """Main consolidation process"""
        try:
            self.setup_repo()
            
            # Get all branches
            branches = self.get_all_branches()
            total = len(branches)
            
            print(f"\n{'='*50}")
            print(f"Found {total} branches to process")
            print(f"{'='*50}\n")
            
            # Switch to main
            print("Switching to main branch...")
            self.run_cmd(["git", "checkout", self.main_branch])
            self.run_cmd(["git", "pull", "origin", self.main_branch])
            
            # Process each branch
            for i, branch in enumerate(branches, 1):
                print(f"[{i}/{total}]", end=" ")
                
                if self.merge_branch(branch):
                    self.merged_count += 1
                else:
                    self.conflict_count += 1
            
            # Push to remote
            print(f"\n{'='*50}")
            print("🚀 Pushing consolidated main branch...")
            self.run_cmd(["git", "push", "origin", self.main_branch, "--force"])
            
            # Delete all remote branches
            print("🗑️  Deleting remote branches...")
            branches_to_delete = self.get_all_branches()
            
            for branch in branches_to_delete:
                try:
                    self.run_cmd(["git", "push", "origin", "--delete", branch])
                    print(f"   ✓ Deleted: {branch}")
                except:
                    print(f"   ✗ Failed to delete: {branch}")
            
            # Print summary
            print(f"\n{'='*50}")
            print("📊 Consolidation Summary")
            print(f"{'='*50}")
            print(f"Total branches processed: {total}")
            print(f"Successfully merged: {self.merged_count}")
            print(f"Conflicts resolved: {self.conflict_count}")
            print(f"Skipped: {self.skipped_count}")
            print(f"\n✨ Consolidation complete!")
            print(f"{'='*50}\n")
            
            # Cleanup
            os.chdir("..")
            self.run_cmd(["rm", "-rf", self.work_dir])
            print("✅ Cleanup done!")
            
        except Exception as e:
            print(f"\n❌ Fatal error: {e}")
            sys.exit(1)

def main():
    consolidator = BranchConsolidator("Cylae", "DefFileGenerator")
    consolidator.consolidate()

if __name__ == "__main__":
    main()
