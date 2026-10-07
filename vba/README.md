# VBA + VS Code + AI Agent Workflow

This setup removes day-to-day copy/paste for VBA changes.

## What you get

- All VBA modules in plain text files under `vba/src`
- AI agent can edit those files directly
- Two macros keep workbook VBA and files in sync

## One-time setup

1. Open `GuinARK_2026_v1.1.xlsm` in Excel.
2. Enable trust access:
   - Windows: Developer -> Macro Security -> Trust access to the VBA project object model
   - macOS: Excel -> Settings/Preferences -> Security (or Security & Privacy), then enable VBA project access if shown
3. Open VBA Editor (Option + F11).
4. Import module file `vba/src/SyncVBA.bas`.
5. Save workbook as macro-enabled (`.xlsm`).

## If you get error H80004005 on import/export

- Why it happens:
  - Excel cannot access the VBA project programmatically.
  - This is often blocked by trust settings, file location permissions, or project protection.
- Fix checklist (macOS):
  - Move workbook to a local path first (for example, Documents), open it there, and retry.
  - Check Excel security settings for VBA project access and enable it if present.
  - Make sure the VBA project is not password protected.
  - In macOS System Settings -> Privacy & Security, allow Excel file access if prompted.
  - Re-open Excel after changing security settings.
- Note:
  - Some Excel for Mac versions restrict VBProject automation more than Windows. If blocked in your build, this workflow may require Windows Excel for full automatic import/export.

## If path shows as https://... on import/export

If Excel opens the workbook via SharePoint URL, `ThisWorkbook.Path` is not a local filesystem path, so sync cannot find folders.

`SyncVBA.bas` is configured to use this machine's project folder:

```text
/Users/guin/Developer/guinark
```

The two path constants have also been updated inside the current workbook and
saved. Excel's embedded `SyncVBA.GetExportPath` was executed and returned
`/Users/guin/Developer/guinark/vba/src`. All other VBA modules and worksheet
cell contents were checked against the pre-repair backup and are unchanged.
The recovery copy is in `backups/sync-path-repair-20261007/`.

Updating the `.bas` file on disk alone does not change the workbook's embedded
macro. If the project moves again, update the embedded constants as well.
Do not run the bulk importer just to update a path: it replaces all imported
VBA modules. Do not export over unimported source edits merely to test the path.

## Daily workflow

1. In Excel, run macro: `ExportVbaProjectToFiles`
2. In VS Code, edit files in `vba/src` (manually or with AI agent)
3. In Excel, run macro: `ImportVbaFilesToProject`
4. Test in Excel and save workbook

## v1.1: migrating Ressourcer to the F27 layout

The migration has been applied and saved in `GuinARK_2026_v1.1.xlsm`.
The original backup is
`backups/GuinARK_2026_v1.1.before-resource-layout.xlsm`.
An additional snapshot taken immediately before the successful migration is
`backups/GuinARK_2026_v1.1.before-resource-layout-current.xlsm`.

The initial layout migration retained the old Ressourcer allocations.
Following clarification, Ressourcer now contains F27's actual period,
allocations, capacity percentages, course inputs and formulas at the same
cell addresses. `F27!A1:AE69` was copied directly to `Ressourcer!A1:AE69`;
all 2,139 cells were checked for matching formulas/constants and calculated
values. This corrected 136 differing cells without regenerating schedules.
A backup including the user's latest edits before this correction is saved
as `backups/GuinARK_2026_v1.1.before-F27-values-20261006-1248.xlsm`.

The migration is complete. F27 has been deleted from the active workbook.
`MigrateResourceLayout`, its private remapping/formula helpers, and the
temporary-tool removal routine have been removed from embedded and source VBA.
Recovery copies retain the historical migration.

`modResourceLayout` retains `AuditResourceDependencies` and
`VerifyResourceLookups`. Its `ApplyResourceSubjectColors` function matches
Ressourcer column A's subject names against Data column C and copies each
matching code's fill from Data column D. The 21 matching subject labels now
use this palette. Values, borders, and unrelated administrative labels are
not changed. Rerun this function if the Data palette changes, then save.

Runtime verification with `modResourceLayout.VerifyResourceLookups()` passed
for all four semester teacher lists and allocated-hour formulas, including
the existing 15-teacher capacity. Verification uses a temporary worksheet
that is deleted afterwards. It does not fill or modify the existing schedules.

Resource references in the new layout:

| Purpose | Range |
| --- | --- |
| Teacher headers | `Ressourcer!B1:X1` |
| Semester 1 | `Ressourcer!B7:X10` |
| Semester 2 | `Ressourcer!B12:X15` |
| Semester 3 | `Ressourcer!B17:X20` |
| Semester 4 | `Ressourcer!B23:X25` |

The source semester module uses these new ranges. Its malformed quote in the
big-layout counting formula has also been corrected so that this edited source
does not introduce a VBA syntax error when imported.

## Important notes

Keep exactly one `SubjectCells` module in the workbook. An identical
`SubjectCells1` duplicate was removed after it caused an ambiguous-name compile
error in `Workbook_SheetCalculate`. Use a clean staging folder for native
module imports; stale exports must not be included in later installations.
After importing, compile the entire VBA project and test calculation with
events enabled.

Semester generation identifies the copied worksheet by comparing sheet names
before and after the copy. It must not use `ActiveSheet` or assume the copy is
the final item in `Worksheets`: on Mac those can still identify an existing
worksheet. Generation suppresses calculation events while building schedules,
explicitly makes each new semester visible without unhiding its template,
restores application settings on success or failure, and freezes only columns
A:D on semester sheets. Overblik freezes columns A:B and rows 1:5.
Overview generation orders the visible main tabs as Ressourcer, Data, Overblik,
Dobbeltbooking, then semesters in ascending order. Hidden sheets remain hidden.
Semester headers include the room from Data's merged G:J room entries:
G47, G49, G51, G53, G55, G57 and G59 for semesters 1 through 7.
For example, a large layout shows `2. Sem A / GBG.A367` and
`2. Sem B / GBG.A367`; a small layout shows `2. Sem / GBG.A367`.
The merged semester-and-room headers use a 22-point font.
Blank rooms leave the original label without a slash; placeholders such as
`???` are shown as entered. `RefreshSemesterRoomHeaders` updates existing
headers without regenerating schedules. Sheet names and Overview labels
remain unchanged.
`VerifySemesterGeneration` runs Plot and then checks the generated Overview;
it returns an explicit error with the generation stage if Plot fails.
It is not a read-only check: use it only when regenerating schedules is intended.

### Semester frozen headers

The four current semester sheets freeze only columns A:D (at E1).
`TemplateBig` and `TemplateSmall` have the same saved pane settings.
Before freezing, clear existing freeze/split settings and reset the window's
scroll row and column to 1, then select E1 and freeze. No semester header rows
are frozen; vertical scrolling is unrestricted. Only Overblik retains its
frozen top headers.
`MakeSemesters` copies these templates, so newly generated semester sheets
inherit the frozen columns without adding VBA logic. Saved pane settings were
verified for all four semesters and both templates, and vertical scrolling was
tested while columns A:D remained visible. Overblik's view is unchanged.

### Overblik blocked days

Reset (`DeleteAllSemesters`) deletes the generated semester sheets and
Overblik. Plot (`Generate_Semesters`) generates and fills the semester sheets
first, then calls `GatherSchedulesToOverblik` to build a fresh overview.
Data, Ressourcer, templates and other input sheets are not cleared.

Overblik mirrors the semester sheets rather than maintaining independent
blocked-date or Wednesday restrictions. The current overview blocks map as follows:

| Overblik rows | Semester source |
| --- | --- |
| 6:20 | `1. Sem!E7:DE21` |
| 22:36 | `2. Sem!E7:DE21` |
| 38:52 | `2. Sem!E28:DE42` |
| 54:68 | `3. Sem!E7:DE21` |
| 70:84 | `4. Sem!E7:DE21` |
| 86:100 | `4. Sem!E28:DE42` |

Columns C:DC correspond to semester columns E:DE. Each source block is copied
as a range, preserving its cell formatting, lunch patterns and subject
conditional formatting. Overview values are linked to the original semester
cells, so schedule edits continue to appear in the overview. Plot refreshes
formatting from the semester sheets; it does not add independent red rules.
The blank top separator is not a blocked-date marker.

`modOverblik` replaces the old flattening implementation. It builds a new
sheet before replacing Overblik. Dobbeltbooking's overview references use
INDIRECT so deleting/recreating Overblik does not turn them into permanent
`#REF!` references. Reset still deliberately clears generated schedules.

`VerifyGeneratedOverblik` checks that generated label cells are merged and
vertical, as well as comparing overview values with their semester sources.
The current saved workbook already has the merged label layout. To make future
Plot runs keep it, replace only the embedded `modOverblik` with
`vba/src/modOverblik.bas`; do not run the bulk importer for this one-module fix.

Overblik is read-only: all cells are locked and worksheet protection is
applied to the current overview and after every Plot rebuild. Users can
select and copy cells but cannot edit their values or formatting.
`ProtectOverblik` deletes copied data validation before locking the overview,
so no dropdown arrows or validation prompts suggest that it is editable.
This applies to the saved overview and future Plot rebuilds; semester
dropdown validation remains intact.
Linked formulas still recalculate when semester entries change. Protection has no
password; it prevents accidental edits rather than providing access security.
Reset can still delete the protected sheet and Plot creates a new protected
overview.

The original Overblik header design is stored in the very-hidden
`TemplateOverblikHeader` sheet. Plot copies its first four rows, merged week
headings, row heights and column widths before linking dates to the semester
sheet. Reset leaves this template intact, so rebuilding does not replace the
original header styling.
The header template's start-date formula is `=Data!M2`, referencing this
workbook only. Its obsolete link to the pre-reset backup was removed.
Excel's Excel/OLE link lists and the saved workbook's external relationships,
formulas and defined names were checked: no external workbook links remain.
The template retains column A's large vertical semester-label formatting.
Each generated overview block applies that formatting and explicitly merges
its 15-cell label range before writing the semester name.
Unused overview cells use the semester sheets' grey separator fill, while
the copied schedule blocks retain their source formatting.

### Updating only `modOverblik`

The bulk importer replaces every imported VBA module, which is not appropriate
for this fix. In Excel's VBA editor, export the existing `modOverblik` as a
backup, remove that module, then import `vba/src/modOverblik.bas`. Compile the
VBA project, run Plot, and run `VerifyGeneratedOverblik`. The verifier must
report the overview values match and will stop if any label is not merged and
vertical. Save the workbook only after those checks pass.

- This sync handles:
  - Standard modules (`.bas`)
  - Class modules (`.cls`)
  - UserForms (`.frm` + `.frx`)
- Existing workflow note:
  - If you already manually exported modules to `vba`, this setup uses that folder directly.
- It does not overwrite worksheet/document modules (like `Sheet1`, `ThisWorkbook`).
- If you need those in source control too, they can be added with an advanced variant.

## Recommended VS Code extensions

- VBA support extension for syntax highlighting
- GitHub Copilot + Copilot Chat for AI edits

## Suggested Git flow

1. Export VBA to files
2. Commit text files in `vba/src`
3. Let AI propose edits with full diff visibility
4. Import back into workbook and test

This gives you reliable diff history and removes repeated manual copy/paste of macro code.
