# GuinARK 2026

Excel-based resource planning workbook with VBA source maintained alongside it.

## Project contents

- `GuinARK_2026_v1.1.xlsm` — current macro-enabled workbook.
- `vba/` — exported VBA modules and the VBA synchronization guide.
- `backups/` and `archives/` — local recovery copies; excluded from Git.

## VBA workflow

1. Read [`vba/README.md`](vba/README.md) before synchronizing source and embedded VBA.
2. Make VBA changes in `vba/src/`.
3. Import/export using the documented workbook macros.
4. Test changes in Excel and save the workbook.

The `.xlsm` workbook is a binary file, so Git cannot show meaningful cell- or
macro-level diffs. Review workbook changes in Excel before committing.

## Git

Backups, archives, Excel lock files, and macOS metadata are excluded by
`.gitignore`. The active workbook is currently under GitHub's 100 MB per-file
limit; Git LFS is not needed at its current size. Check `git status` and review
staged workbook changes before committing or pushing.
