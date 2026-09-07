---
name: design-doc-formatting-consistency
description: Scan a design-doc workbook for stray font-size and cell-merge irregularities — a cell whose font size or merge span breaks from the surrounding pattern, which usually signals a leftover copy-paste artifact, an incompletely-applied edit, or a duplicated header block that never got updated. This is a formatting/editing-hygiene check, distinct from design-doc-internal-consistency (cross-reference completeness), design-doc-io-table-check (I/O table completeness), xlsx-db-column-check (DB column existence), and design-doc-typo-check (actual Japanese-language proofreading). Use as part of a full program REV, or when the user specifically asks about フォントサイズ/セル結合のブレ. When the user asks to REV a single design-doc workbook (not scoped to a whole WG folder), run this together with its 5 sibling single-program skills — `design-doc-internal-consistency`, `design-doc-io-table-check`, `xlsx-db-column-check`, `naming-standard-compliance`, `design-doc-typo-check` — as one combined pass, not standalone.
---

# design-doc-formatting-consistency

Finds stray font-size and cell-merge irregularities across every sheet of one program's design-doc
workbook — see the frontmatter `description` above for what these irregularities typically signal.
Unlike the other REV skills, this one isn't grounded in the project's checklist or a cross-reference
rule; it's a pure structural-hygiene sweep, worth surfacing to the designer even though these aren't
"wrong" in the same sense as a broken cross-reference.

**When the user asks to REV a single program's design-doc workbook, run all 6 single-program-scoped
skills together in that one pass** — this skill, `design-doc-internal-consistency`,
`design-doc-io-table-check`, `xlsx-db-column-check`, `naming-standard-compliance`, and
`design-doc-typo-check` — and report their findings combined as one result, not as separate reports
the user has to request individually. When these are run as parallel background agents, do not post
a status update each time one agent finishes — wait until every agent in the batch has completed,
then compose and post the single combined report in one message. Also dump the target workbook's
sheets once yourself before launching the batch, and point each agent at those scratchpad text files
instead of letting every agent re-dump the same workbook independently — see
`_shared/xlsx-excel-com-dump.md`'s "dump once, share the text" section (this skill's own font-size/
merge scans still need live Excel COM access per-agent — only the bulk text dump is shared). This
does not extend to the
WG-folder-scoped comparison skills (`db-design-cross-consistency`, `cross-function-similarity-check`)
— bundle those in only when the user's request is itself folder-scoped, or they explicitly ask for
them.

## Environment

Read `~/.claude/skills/_shared/xlsx-excel-com-dump.md` first for the core dump/COM technique, then
read `~/.claude/skills/_shared/xlsx-formatting-scan.md` — it has the exact font-size/cell-merge
detection technique (including the DBNull guard for mixed-formatting cells) and, critically, the
noise-filtering rules learned from running this on real workbooks. **Read that second file fully
before running anything here** — a naive sheet-wide "most common size/span" comparison for merges
produces an unusable flood of false positives; it explains why and what to do instead (a local,
row-proximity comparison for merges; a set of known systematic patterns to exclude for font sizes).
(This technique lives in its own file, separate from the main shared dump doc, specifically because
no other REV skill needs it — don't merge it back.)

**Don't dump or scan `詳細設計書` sheets — this is a deliberate coverage-for-tokens trade-off, not a
claim that the sheet is always empty.** Unlike `design-doc-internal-consistency`/
`naming-standard-compliance` (which skip it because their specific checks genuinely don't need
detail-design content), this skill's blanket per-cell scan *can* find real things there — a real
run against a sub-program (SXJCB147, a ﾌﾞﾌﾟﾛ) found two genuine font-size irregularities inside its
`詳細設計書` sheet (in its ﾊﾟｯｹｰｼﾞ構成 and 引数一覧 sections, which were populated, not
header-only). The user explicitly chose to exclude it anyway for the token savings, accepting that
findings of that kind will be missed going forward. If the user asks for a more thorough pass, or
specifically asks about 詳細設計書 formatting, include it then.

## Procedure

### 1. Font-size scan

For every worksheet in the target workbook **except `詳細設計書`** (see above), loop its non-empty
cells (reuse the coordinates you'd
already get from the standard bulk `Value2` dump) and record each cell's `Font.Size`. Per sheet,
tally the size frequencies and take the mode. List every cell whose size differs from that sheet's
mode — this raw list will be large and mostly noise.

**Before reporting anything, filter out the deliberate systematic patterns** documented in the
shared doc (revision-history annotation columns, section-title/header template cells, control-matrix
column-group headers, and anything on a UI-mockup/screenshot sheet). What's left after that filter —
genuinely isolated cells, especially ones sitting inside an otherwise-uniform repeating list — is
what's worth reporting.

### 2. Cell-merge scan

For the same non-empty cells, check `MergeCells`; when true and the cell is the merge's own anchor,
record the merge's row/column span. **Do not compare spans sheet-wide by column position** — compare
locally: for merges at the same column, flag pairs whose spans differ while sitting within a handful
of rows of each other (see the shared doc for why and the exact approach). This surfaces things like
a duplicated header block whose right-side copy still carries stale content, or one row of a
repeating item list that's merged differently than every sibling row around it.

After generating candidates, do a manual pass to drop remaining template-boilerplate coincidences
(e.g. a "No." column naturally spanning differently right next to a "特記事項" footer row) — only
report pairs that plausibly represent an actual duplicated block or list member, not just two
adjacent structural cells that happen to serve different roles.

### 3. Cross-check anything suspicious against content, not just shape

A merge or font-size anomaly is a **lead**, not a finding by itself — when something survives the
filtering above, look at what's actually written in the flagged cell(s) before reporting. A
duplicated header block with a mismatched merge span is far more actionable once you've also checked
whether its *content* is stale (e.g. an old system name, a blank date, a wrong ID) — that's often
the real, citable defect, with the formatting anomaly serving as how you found it in the first place.

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on. Write in the user's language.

For each real finding: cite the exact sheet name and cell coordinate(s), state what's different (the
size/span and what the surrounding pattern is), and — when you cross-checked content per step 3 —
say what that content difference actually is, since that's usually the more useful half of the
finding. Group by sheet only when several findings share one.

Omit entirely: the raw per-sheet size/merge dump, a tally of how many cells were checked, and any
finding that turned out to be one of the documented systematic patterns (revision-history columns,
template headers, mockup sheets, control-matrix group headers) — those are not worth even a one-line
mention once identified as such, since they're expected and consistent by design. If a whole class of
irregularity couldn't be checked (e.g. a sheet's used range was too large to scan in full), say so
briefly — that's a coverage gap worth flagging, not a process detail.
