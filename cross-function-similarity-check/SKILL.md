---
name: cross-function-similarity-check
description: Compare similar processing across multiple programs' 機能定義書/画面設計書/更新条件表 within a WG folder — e.g. several programs each fetching from the same master table, or several update-condition tables deriving the same common column — and flag cases where equivalent logic diverges (a missing exclusion condition, a different operator, a missing column, a different source formula for a common column) as suspected design mistakes. Use when the user asks to compare 機能間 for 似た処理/似た取得 and wants outliers surfaced, as distinct from `design-doc-internal-consistency` (checks one program's internal traceability) or `db-design-cross-consistency` (checks DB column *definitions*, not usage).
---

# cross-function-similarity-check

Finds groups of "should behave the same" processing scattered across different programs' design
docs in one WG, and flags the member that doesn't match the rest of its group. This is a
**behavioral/usage** comparison (what each program actually fetches/filters/derives), not a
structural one — for DB column definition consistency use `db-design-cross-consistency`, for one
program's own I/O-table declaration completeness use `design-doc-io-table-check`, for its other
internal cross-references use `design-doc-internal-consistency`.

The underlying assumption: when N programs in the same WG all read the same master table (or all
populate the same common column), their filter/exclusion conditions, fetched columns, sort order,
and value-derivation formulas should usually agree. A program that's the odd one out — missing a
condition every sibling has, using a different comparison operator, omitting a column every sibling
fetches, or deriving a common column from a different source — is a strong candidate for a design
mistake (most often: a condition that should have been copied from a similar program's screen but
wasn't). This is a **heuristic, candidate-generating** check, not a proof of a bug — some divergence
is legitimate business difference. Always report findings as "worth confirming," tiered by
confidence, never as confirmed defects.

## Environment

Read `~/.claude/skills/_shared/xlsx-excel-com-dump.md` first for the PowerShell + Excel COM dump
script (no Python/Node available here), and its "Excluding struck-through / grayed-out rows from
review" section. Struck-through/grayed rows are deleted-in-spirit — run the targeted formatting
scan against the label column (項目名/画面項目名) of every section you extract from, **and against
the table-name column of every 画面設計書 "参照ｴﾝﾃｨﾃｨ" block a unit's table-set/column references
come from** (a struck-through block invalidates every 取得項目/検索条件/結合条件 line pulled from
it, not just the table-name line), and drop flagged rows/blocks before clustering. Watch in
particular for `処理区分="..."の場合` branch variants — this project commonly deprecates one whole
branch (spanning several numbered sub-sections) while leaving the sibling branch active, and citing
a unit from the deprecated branch as this program's real behavior produces exactly the kind of false
outlier this skill exists to avoid. This matters even more here than in single-program checks: a
leftover struck-through condition copied from an older program is exactly the kind of noise that
would produce a false "this program has an extra condition the others lack" finding.

## Procedure

### 1. Enumerate scope

**If the user invoking this skill hasn't already named a target WG folder (or a narrower subfolder)
in their request, stop and ask before doing anything else** — don't guess a folder or search the
whole project. This check runs across an entire WG's worth of workbooks and is expensive to run
against the wrong scope. Once a folder is confirmed, find every 機能定義書/画面設計書 workbook in
scope (`find ... -iname "*.xlsx"`, excluding 表紙-only or テーブルレイアウト files). This can be
dozens of files — batch the dump through one PowerShell invocation (one `$excel` instance, loop
over paths), not one call per file.

### 2. Extract a "processing unit" per program

For each workbook, dump and read these sections, and record one **unit** per numbered subsection /
per table:

- **画面設計書 "Ⅲ．画面表示仕様"**: each numbered subsection (e.g. "(3)加工手順",
  "(4)出力実績表項目取得") is one unit. Record: program ID, subsection title, the "参照ｴﾝﾃｨﾃｨ" block
  (table IDs + aliases, ignore Y=画面/Z=ﾛｸﾞｲﾝ情報), 取得項目 (fetched `<alias>.<column>` list),
  検索条件 (`<alias>.<column> <op> <value>` list), 結合条件 (join conditions), ｿｰﾄ順.
- **機能定義書 "Ⅳ．機能処理概要"**: each described step is a lighter-weight unit — record the step
  text and which table(s)/keywords it mentions (used only for prose-level clustering, see step 3).
- **更新条件表(<TableID>) sheets**: each row is a unit keyed by the *column being set*, not the
  table — record 項目名 and its "取得内容" (the source formula/expression for that column's value),
  across every 更新条件表 sheet in scope project-wide, not just per program.

### 3. Cluster units across programs

Clustering is heuristic — use judgment, don't demand exact string matches:

- **Primary key — table-set fetched.** Normalize each Ⅲ．画面表示仕様 unit's 参照ｴﾝﾃｨﾃｨ table-ID set
  (sorted, ignoring aliases). Units with an identical or near-identical table-set (e.g. same single
  master table, possibly joined to one extra lookup in only some units) form a cluster — these are
  the strongest, lowest-noise clusters since they're reading literally the same data source.
- **Secondary key — purpose keyword.** For units that don't share a table-set, tokenize the
  subsection title (strip the leading "(N)") and compare against other units' titles/keywords (e.g.
  every unit whose title contains "得意先" + "取得", regardless of exact table joined) — these
  clusters are noisier (the same *business* lookup implemented against different table shapes) but
  still worth comparing at the condition level (e.g. "does everyone exclude inactive records the
  same way").
- **Common-column key — 更新条件表 rows.** Separately, group all 更新条件表 rows project-wide by
  項目名 (e.g. every 会社コード, 削除フラグ, 部門GRP row across every table in scope) and compare
  their 取得内容 formulas within each group. This mirrors `db-design-cross-consistency`'s
  same-name-column check but compares *how the value is derived*, not the column's physical
  type/length.

Discard clusters of size 1 (nothing to compare against) and clusters so large/generic they're
meaningless (e.g. every table has a 更新日時 row deriving from `SYSDATE` — not interesting; focus on
columns whose derivation plausibly varies, like 会社コード, 部門GRP, or business keys).

### 4. Diff within each cluster and flag outliers

Within each cluster of 2+ members, compare:

- **Missing condition.** A 検索条件/結合条件 that appears in most members but is absent in one —
  classic candidates: a status/exclusion flag (削除フラグ, 有効フラグ, 状態区分), a company/tenant
  scoping condition (会社コード = ログイン情報の会社コード), a date-range bound. Flag high-confidence
  if the missing condition looks like a safety/scoping check; flag lower-confidence if it looks like
  an optional business filter.
- **Different operator for an apparently-equivalent condition** (e.g. one unit uses `>=` and another
  `>` against what should be the same date-boundary semantics).
- **Missing fetched column.** A 取得項目 column that every other member of the cluster fetches but
  one doesn't (e.g. everyone else also fetches 表示順/名称 alongside the code, one unit fetches only
  the code) — worth flagging since the odd-one-out screen may display an incomplete row.
- **Sort-order divergence** where the units otherwise look like "the same kind of list screen" —
  lower confidence, note it rather than asserting a defect.
- **Divergent 取得内容 formula for a common column** (e.g. 会社コード populated from ログイン情報 in
  every 更新条件表 except one, where it's copied from a joined table's own column instead) — flag
  high-confidence, this is a common source of company/tenant-scoping bugs.

### 5. Reporting

The output goes to the designers who own these programs, not into your own working notes — report
only what they'd actually need to act on.

Group findings by cluster (name each cluster by its shared table-set or keyword), but only include a
cluster that actually produced a finding — don't list clusters that compared clean. For each
flagged cluster: list the outlier and the sibling(s) it diverges from with a clickable
`[filename](relative/path)` link and the exact sheet/subsection/cell cited, then show the outlier
side-by-side against the majority pattern (a small comparison table, not just prose). Order
high-confidence findings before low-confidence ones. Tag every finding **確度: 高/低**:
- 高 — the divergent member is missing something every sibling has and the missing item looks like
  a safety/scoping/completeness check.
- 低 — plausible legitimate business difference; surfaced for the human to confirm, not asserted as
  a bug.

Omit entirely: a tally of workbooks scanned / units extracted / clusters formed. The one exception
worth a one-line mention: a cluster you had to discard as too noisy to compare, or a program whose
workbook couldn't be read — that's a coverage gap the designers should know the sweep didn't
actually cover, not a process detail.
