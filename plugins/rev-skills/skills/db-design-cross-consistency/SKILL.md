---
name: db-design-cross-consistency
description: Cross-check DB design (テーブルレイアウト) Excel files against each other and against the project's common-column/ID-naming standards — common audit columns (会社コード/部門GRP/登録者/更新日時/排他フラグ等) must be identical everywhere, same-named key columns referenced from multiple tables (品目コード/工程コードなど) should share the same physical type/length, and every table must be registered in the DB一覧. Use when the user asks to REV/レビュー a WG's DB設計書 folder as a whole, or asks whether tables are "整合しているか" / "統一されているか" across a folder.
---

# db-design-cross-consistency

Cross-checks multiple ﾃｰﾌﾞﾙﾚｲｱｳﾄ workbooks against each other (not against a program's usage — for
that, see `xlsx-db-column-check` for column-level usage or `design-doc-io-table-check` for whether a
table is declared as used at all). Grounded in this project's own common-design rules:

- `01_Doc/04_共通設計/05.ｼｽﾃﾑ共通設計書.xlsx`, sheet **"更新条件共通項目"** — defines the standard
  audit-column block every table must have (会社ｺｰﾄﾞ, 登録者, 登録日時, 更新者, 更新日時,
  更新ﾎｽﾄ名, 更新ﾌﾟﾛｸﾞﾗﾑID, 排他ﾌﾗｸﾞ, 部門GRP — plus 処理区分 for 履歴 tables). Read this sheet
  first each time; it is the authoritative baseline.
- `01_Doc/04_共通設計/05.ｼｽﾃﾑ共通設計書.xlsx`, sheet **"各ID採番"** — table ID naming rule:
  `①[T=ﾃｰﾌﾞﾙ|V=ﾋﾞｭｰ] + ②<3-char JOBｺｰﾄﾞ> + ③[M=ﾏｽﾀ|D=ﾃﾞｰﾀ|V=ﾄﾗﾝ|A=累積|P=ﾊﾟﾗﾒｰﾀ] + ④<3-digit
  zero-padded seq>` (e.g. `TXJAM025`=T/XJA/M/025, `TSJCV501`=T/SJC/V/501,
  `VSJAM060`=V/SJA/M/060). File IDs (`F...`) use a separate 2-part rule (`F + JOBｺｰﾄﾞ + seq`, no
  data-kind letter).
- `01_Doc/06_システム設計書（一覧、管理台帳）/06-06_DB一覧_共通.xlsx` — the master registry every
  table should appear in exactly once.

## Environment

Read `~/.claude/skills/_shared/xlsx-excel-com-dump.md` first for how to dump `.xlsx` sheets to text
via PowerShell + Excel COM (no Python/Node available here) — including its "Excluding
struck-through / grayed-out rows from review" section. A ﾃｰﾌﾞﾙﾚｲｱｳﾄ row an author struck through or
grayed out is a deleted-in-spirit column, not a live one — run that section's targeted formatting
scan against the 項目名 column of each table before feeding its rows into the common-column or
key-column comparisons in step 4, and drop any flagged row.

## Procedure

### 1. Establish the baseline

Dump and read "更新条件共通項目" and "各ID採番" from `05.ｼｽﾃﾑ共通設計書.xlsx` (these rarely change,
so it's fine to re-read them at the start of every run rather than assume a prior session's
memory of them is still accurate).

### 2. Enumerate the target scope

**If the user invoking this skill hasn't already named a target WG folder (or a narrower subfolder)
in their request, stop and ask before doing anything else** — don't guess a folder or search the
whole project. This check runs across an entire WG's worth of workbooks and is expensive to run
against the wrong scope.

Once a folder is confirmed (e.g. `<WG>/07_データベース・ファイル設計書(仮)`, optionally scoped
further to a subfolder like `PH3`) — **note this is a TOP-LEVEL project folder named
`<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)` (e.g. `11_工程管理WG\07_データベース・
ファイル設計書(仮)`), a sibling of `01_Doc`, NOT nested inside it**, even though most other
design-doc types live under `01_Doc\...` — `find` every `*.xlsx`/`*.xlsm` in scope, excluding
`WF`-suffixed workbooks only if the user wants base tables only (WF = workflow-draft shadow tables
and are usually structurally identical to their base table — worth including if checking
common-column consistency, since they should match too).

**The same table ID can legitimately appear more than once in scope** — once directly under the WG
folder and again under a `PH2`/`PH3` subfolder. Group discovered files by table ID; for any ID with
more than one file, resolve to exactly one by phase precedence — **PH3 > PH2 > the WG folder's top
level** — and drop the rest before running the comparisons below. Don't feed more than one copy of
the same table ID in as if they were independent tables (that would either double-count the table or
silently compare against a superseded copy). This phase-precedence rule replaced an earlier, less
reliable attempt at this same problem based on comparing file-modified timestamps, after a `PH2`
copy of `TXJCM058` turned out to be the current, more complete layout despite the top-level copy
looking no different at a glance — a real false-positive this produced in a past review.

### 3. Dump each table's ﾃｰﾌﾞﾙﾚｲｱｳﾄ sheet

Use the **"Batch variant"** script from `_shared/xlsx-excel-com-dump.md`'s cross-session cache
section, with every resolved file's `OnlySheetPatterns = @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ")` — this skips 改訂履歴/SQL/JAG*
sheets (historical/generated, not the source of truth), and since these same table-layout files get
read repeatedly across many different programs' single-program REVs (`xlsx-db-column-check`) as well
as this WG-wide audit, a table already dumped by either one is a cache hit for the other. This can be
dozens of files, and the batch script checks every file's cache validity before ever launching
Excel — it opens a single shared `Excel.Application` only for the files that are genuine misses (a
cold WG-wide run will still redump most of them, but a repeat run, or one after several
single-program REVs have already warmed the cache, can launch Excel zero times). Build the file list
within your own turn rather than spawning sub-agents per file (see the shared dump doc's note on not
fanning out for large table/sheet lists) — the batch script itself already handles looping over many
files in one pass.

### 4. Checks

**A. Common audit-column consistency.** For every table, extract the rows matching the 9 common
column names from step 1's baseline. Compare each column's 属性 (type) and 桁数 (length) across all
tables in scope:
- Flag any table where a common column's type/length **deviates** from the majority/standard (e.g.
  everyone else has `会社ｺｰﾄﾞ NVARCHAR2(5)` but one table has `NVARCHAR2(10)`).
- Flag any table **missing** one of the 9 common columns entirely (excluding legitimate cases like
  a pure トラン/検索条件 table that intentionally has a reduced audit set — use judgment, and check
  whether the table name suggests a 履歴 table, which should additionally have 処理区分 per the
  common-design's 2nd rule).
- Also check column **order**: the common columns are conventionally the first ~9-10 physical
  rows in every ﾃｰﾌﾞﾙﾚｲｱｳﾄ; a table where they're reordered or interleaved with business columns is
  worth flagging even if the types match, since it deviates from every other table in the project.

**B. Same-name-different-shape key columns.** Across all tables in scope, group columns by 項目名
(Japanese name) — e.g. every table with a `品目ｺｰﾄﾞ`, `工程ｺｰﾄﾞ`, `加工手順ｺｰﾄﾞ`, `部門GRP`, `事業部ｺｰﾄﾞ`
column. For each group, compare 属性/桁数 across all occurrences:
- Flag any column whose type/length disagrees with the rest of the group — this usually means a
  join between two tables will silently truncate or mismatch (e.g. a master table defines
  `品目ｺｰﾄﾞ NVARCHAR2(60)` but a child table that's supposed to reference it only allocated
  `NVARCHAR2(30)`).
- This check is naturally noisy for generic short names (`No`, `名称`, `区分`) — focus on columns
  that look like keys/codes (`...ｺｰﾄﾞ`, `...ID`, `...GRP`), not every same-named column in the
  corpus.

**C. Table ID naming-rule compliance.** For each table's ID (from ﾃｰﾌﾞﾙﾚｲｱｳﾄ row 6), check it parses
against the rule in step 1: object-type letter matches whether it's actually a table or a view
(cross-check against the sheet's own "ﾃｰﾌﾞﾙ名" — does it say ビュー anywhere, or does its ﾃｰﾌﾞﾙﾚｲｱｳﾄ
look like a SELECT/JOIN definition rather than a stored table?), JOBコード matches the WG's actual
job codes in scope, and the data-kind letter (M/D/V/A/P) is a plausible match for the table's name
(e.g. a name ending in マスタ should be `M`, ...履歴/...ﾄﾗﾝ should usually be `V` or `A`). Flag clear
mismatches; treat borderline ones (e.g. M vs D for an ambiguous name) as a lower-confidence note.

**D. DB一覧 registration.** Check every table ID found in scope appears in
`01_Doc/06_システム設計書（一覧、管理台帳）/06-06_DB一覧_共通.xlsx`, and flag any that don't. Also
flag (lower priority, informational) any 06-06 entries claiming to belong to this WG/folder that
have no corresponding ﾃｰﾌﾞﾙﾚｲｱｳﾄ file in scope — could mean the design file was moved/renamed.

## Reporting

The output goes to the designer who owns these tables, not into your own working notes — report
only what they'd actually need to act on.

Group findings under A/B/C/D headers above, but only include a header if it actually produced a
finding. For A and B, always show a small comparison table (or list) of the conflicting values
across tables, not just "inconsistent" — the designer needs to see what the majority value is vs.
the outlier to judge which side is wrong. Link every table's design file with a relative markdown
link. Order findings by impact within each group — a type/length mismatch on a key column that would
silently truncate data on join belongs before a column-ordering nit. For C, mark borderline
data-kind-letter judgment calls (M vs D) as needing the designer's confirmation, not as confirmed
defects.

Omit entirely: tables/columns that matched the baseline cleanly, and a running tally of how many
tables/columns were checked. The one exception worth a one-line mention: any table in scope whose
ﾃｰﾌﾞﾙﾚｲｱｳﾄ sheet couldn't be read or was missing entirely — that's a coverage gap, not a process
detail.
