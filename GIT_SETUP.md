# Git setup

The Git repository uses `main`. Backups and archives stay local; all VBA
modules live in `vba/src/`.

## Before the first commit

Review the workbook for data that should not be shared. Excel workbooks are
binary in Git; changes cannot be reviewed as normal text diffs. Backups and
archives are intentionally ignored and remain local.

```bash
git status --short
git add .gitignore README.md GIT_SETUP.md vba/
# Add the workbook only after reviewing it for shareable data:
git add GuinARK_2026_v1.1.xlsm
git diff --cached --stat
git diff --cached --name-only
git commit -m "Initial GuinARK project"
```

Create a private GitHub repository after authenticating the GitHub CLI:

```bash
gh auth login --hostname github.com
gh repo create guinark --private --source=. --remote=origin --push
```

The current workbook is below GitHub's 100 MB per-file limit, so Git LFS is
not currently necessary. Consider LFS if the workbook grows substantially.
