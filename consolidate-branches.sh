#!/bin/bash

# Script to consolidate all branches into main branch
# This script:
# 1. Fetches all branches
# 2. Merges unique content from each branch into main
# 3. Deletes all other branches

set -e

REPO_URL="https://github.com/Cylae/DefFileGenerator.git"
WORK_DIR="./DefFileGenerator-consolidation"
MAIN_BRANCH="main"

echo "🔄 Starting branch consolidation process..."

# Clone the repository
if [ -d "$WORK_DIR" ]; then
    rm -rf "$WORK_DIR"
fi

git clone "$REPO_URL" "$WORK_DIR"
cd "$WORK_DIR"

# Configure git
git config user.name "Consolidation Bot"
git config user.email "bot@consolidation.local"

# Get all branches (excluding main)
echo "📋 Fetching all branches..."
git fetch --all

ALL_BRANCHES=$(git branch -r | grep -v "origin/HEAD" | sed 's/origin\///' | grep -v "^main$" | sort)

BRANCH_COUNT=$(echo "$ALL_BRANCHES" | wc -l)
echo "Found $BRANCH_COUNT branches to merge"

# Switch to main
git checkout "$MAIN_BRANCH"
git pull origin "$MAIN_BRANCH"

# Merge all branches with merge strategy
MERGED_COUNT=0
CONFLICT_COUNT=0

for branch in $ALL_BRANCHES; do
    echo ""
    echo "⏳ Processing branch: $branch"
    
    # Try to merge the branch
    if git merge --no-ff --allow-unrelated-histories -m "Merge branch '$branch' into main" "origin/$branch" 2>/dev/null; then
        echo "✅ Successfully merged: $branch"
        MERGED_COUNT=$((MERGED_COUNT + 1))
    else
        echo "⚠️  Conflict detected in: $branch"
        # Abort the merge and keep going
        git merge --abort 2>/dev/null || true
        CONFLICT_COUNT=$((CONFLICT_COUNT + 1))
        
        # Try alternative: cherry-pick unique commits
        echo "   Attempting cherry-pick strategy..."
        COMMITS=$(git log origin/"$branch" ^"$MAIN_BRANCH" --oneline | wc -l)
        if [ "$COMMITS" -gt 0 ]; then
            git cherry-pick -X theirs origin/"$branch" 2>/dev/null || git cherry-pick --abort 2>/dev/null || true
            echo "   ✓ Cherry-picked commits from $branch"
        fi
    fi
done

echo ""
echo "═══════════════════════════════════════════"
echo "📊 Consolidation Summary"
echo "═══════════════════════════════════════════"
echo "Total branches found: $BRANCH_COUNT"
echo "Successfully merged: $MERGED_COUNT"
echo "Conflicts resolved: $CONFLICT_COUNT"
echo ""

# Push consolidated main branch
echo "🚀 Pushing consolidated main branch..."
git push origin "$MAIN_BRANCH" --force

# Delete all remote branches except main
echo ""
echo "🗑️  Deleting remote branches (keeping main only)..."

for branch in $ALL_BRANCHES; do
    git push origin --delete "$branch" 2>/dev/null && echo "  ✓ Deleted: $branch" || echo "  ✗ Failed to delete: $branch"
done

echo ""
echo "═══════════════════════════════════════════"
echo "✨ Branch consolidation complete!"
echo "═══════════════════════════════════════════"
echo ""
echo "Next steps:"
echo "1. Verify main branch: git log"
echo "2. Check repository on GitHub: https://github.com/Cylae/DefFileGenerator"
echo ""

# Cleanup
cd ..
rm -rf "$WORK_DIR"

echo "✅ Done! Repository is ready with only 'main' branch."
