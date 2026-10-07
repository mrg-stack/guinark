# Git setup

The Git repository uses `main`. Backups and archives stay local; all VBA
modules live in `vba/src/`.

## Remote repository

The private repository is https://github.com/mrg-stack/guinark.
`origin` points to it, and local `main` tracks `origin/main`.

On this Mac, `~/.config` is owned by root. GitHub CLI therefore uses a
user-writable configuration directory instead:

```bash
export GH_CONFIG_DIR="$HOME/Library/Application Support/gh"
gh auth status
git push
```

Set this environment variable before authenticated GitHub CLI commands or
HTTPS Git pushes on this machine. Never commit authentication files or tokens.

## Committing updates

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
git commit -m "Describe the changes"
```

Push reviewed commits to the existing private repository:

```bash
export GH_CONFIG_DIR="$HOME/Library/Application Support/gh"
git push
```

The current workbook is below GitHub's 100 MB per-file limit, so Git LFS is
not currently necessary. Consider LFS if the workbook grows substantially.
