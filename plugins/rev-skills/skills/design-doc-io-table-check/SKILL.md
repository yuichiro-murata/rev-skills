---
name: design-doc-io-table-check
description: Check whether a program's 機能定義書「Ⅲ．入出力定義」table (the CRUD list of every DB table/file the program touches) is complete and accurate in both directions — every table actually used anywhere in 画面設計書/更新条件表/帳票設計書 (including via delegation to `07.共通項目取得.xlsx`) is declared with the right C/R/U/D flags, and every declared table is genuinely used and has a real DB一覧 entry + テーブルレイアウト file. This is the single most complex, highest-yield check in a program REV — deep delegation-tracing, WG-folder DB-layout lookups with PH2/PH3 precedence, exact-ID-match discipline, and struck-through-block exclusion all apply — so it is split out from its lighter sibling `design-doc-internal-consistency` (event/response/screen-item-ID/message-ID/three-way-match/exclusive-control checks) to get undivided attention. Use when the user asks to check 入出力定義表/CRUD一覧 completeness, or whether a design doc's declared tables match what it actually uses. When the user asks to REV a single design-doc workbook (not scoped to a whole WG folder), run this together with its 5 sibling single-program skills — `design-doc-internal-consistency`, `xlsx-db-column-check`, `naming-standard-compliance`, `design-doc-formatting-consistency`, `design-doc-typo-check` — as one combined pass, not standalone.
---

# design-doc-io-table-check

Checks one program's 機能定義書「Ⅲ．入出力定義」table for completeness in both directions, grounded
in this project's own checklist: `01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx`
(sheet "ﾁｪｯｸﾘｽﾄ", items 3-2/3-3) — see the frontmatter `description` above for exactly what "complete
in both directions" covers and why this was split out of `design-doc-internal-consistency`.

**When the user asks to REV a single program's design-doc workbook, run all 5 sibling single-program
skills together in that one pass** — `design-doc-internal-consistency`, `xlsx-db-column-check`,
`naming-standard-compliance`, `design-doc-formatting-consistency`, and `design-doc-typo-check` — and
report their findings combined as one result, not as separate reports the user has to request
individually. When these are run as parallel background agents, do not post a status update each
time one agent finishes — wait until every agent in the batch has completed, then compose and post
the single combined report in one message. Also dump the target workbook's sheets once yourself
before launching the batch, and point each agent at those scratchpad text files instead of letting
every agent re-dump the same workbook independently — see `_shared/xlsx-excel-com-dump.md`'s "dump
once, share the text" section. This does not extend to the WG-folder-scoped comparison skills
(`db-design-cross-consistency`, `cross-function-similarity-check`) — bundle those in only when the
user's request is itself folder-scoped, or they explicitly ask for them.

## Environment

Read `~/.claude/skills/_shared/xlsx-excel-com-dump.md` first for how to dump `.xlsx` sheets to text
via PowerShell + Excel COM (no Python/Node available here) — including its "Excluding
struck-through / grayed-out rows from review" section. These docs are often copied from older
workbooks, and rows/items an author struck through or grayed out are meant to be deleted, not
active content. Before diffing, run that section's targeted formatting scan against the item/label
column of each section you're pulling from (画面項目, 項目名, 画面項目名, 参照ｴﾝﾃｨﾃｨ table-name
column, etc.) and drop any flagged row — otherwise leftover crossed-out entries produce false "used
but not declared" findings.

**Don't dump or read `詳細設計書` sheets.** This check draws from 機能定義書/画面設計書/
ﾁｪｯｸ処理設計書/更新条件表/帳票設計書 only. Across every program reviewed so far, 詳細設計書 has
turned out to be a near-empty header-only sheet anyway, so skipping it is a pure token saving with
nothing lost.

## Procedure

**Quick reference — the checks in order (each detailed below with confirmed real-world examples):**

1. 機能定義書「Ⅲ．入出力定義」often has TWO table-ID sources (a live CRUD summary table + a separate
   INPUT/OUTPUT breakdown table right below it) — check the breakdown table's own strikethrough
   formatting before trusting it; it's frequently dead boilerplate.
2. Collect every table ID/alias from: Ⅲ．入出力定義 itself, 画面設計書's 参照ｴﾝﾃｨﾃｨ blocks (or the
   batch/report-program equivalents), AND every 更新条件表 sheet's *body* (not just its header) —
   the 取得内容/取得条件 column often cites other tables by Japanese name only, never by ID, so
   search by name too before calling a table unused.
3. Follow every `※<共通設計書名>.<項目> 参照` delegation line to its actual target sheet (check both
   the WG-specific sheet and the shared sheet in the common-design workbook) — a table only reached
   through a delegated block still counts as used.
4. Before counting any 参照ｴﾝﾃｨﾃｨ block as usage evidence, run the struck-through/grayed-out
   formatting scan on it first — run this unconditionally, on small blocks too, in both diff
   directions (confirming "used" and confirming "not used" both need it).
5. Require an **exact** ID string match throughout (never prefix/substring) — suffix variants
   (`WF`, numeric suffixes like `_31`) are different tables from their base ID, and this applies to
   locating the DB layout file too, not just the diff itself.
6. Diff declared vs. used in both directions, then separately verify each C/R/U/D letter against
   what the live references actually do (a table can be "used" without every declared letter being
   backed by a real operation, or vice versa). Also confirm DB一覧 registration and テーブルレイアウト
   existence — search the correct top-level `<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)`
   folder (PH3 > PH2 > top-level precedence), plus `01_Doc\07_データベース・ファイル設計書\` for
   基準情報(JA)-prefix common tables, before ever reporting a layout as missing.

**Detailed rules and confirmed examples:**

**機能定義書「Ⅲ．入出力定義」routinely contains TWO separate table-ID sources, not one — a CRUD
summary table (columns/rows listing each table ID with its C/R/U/D flags) AND, directly beneath
it, a separate "1．INPUT"/"2．OUTPUT" breakdown table (per-table rows with No./ID/名称/用途
columns) that restates largely the same tables in more prose-like form.** This second, breakdown
table is a well-known trouble spot in this project: it is frequently left ENTIRELY red+struck-through
top to bottom (including its own header row) as leftover, never-cleaned-up boilerplate from the
source workbook it was copied from — while the CRUD summary table right above it in the same
sheet is fully live. Confirmed for real on `PXJCO134_流動停止解除.xlsx`: the OUTPUT breakdown
table at rows 58-64 was entirely red+struck (including a row citing `TXJCM502`/`TXJCA301`, whose
apparent "duplicate No.5" numbering was flagged as a defect by a review that hadn't checked this
block's formatting — a false finding, since the whole block is dead). The same shape (an entirely
struck INPUT/OUTPUT breakdown table, unrelated to the live CRUD summary table above it) was also
the root of a separate finding on `PXJCB102_製造ｵｰﾀﾞｰ完了処理.xlsx`. **Whenever "Ⅲ．入出力定義"
is in scope, check the INPUT/OUTPUT breakdown table's formatting as its own explicit step — don't
assume it's live just because the CRUD summary table above it is, and don't assume a mandate to
"check formatting before citing 参照ｴﾝﾃｨﾃｨ evidence" (elsewhere in this section) already covers
it, since this breakdown table is a different structure in a different sheet (機能定義書 itself,
not 画面設計書) that is easy to mentally file under "not what that rule meant."** If found entirely
struck, exclude it wholesale from every check in this section (don't cite its rows as I/O evidence,
don't flag numbering/content oddities inside it) and just note in the report that it's dead
boilerplate, not a finding to enumerate as a defect.

Collect every table ID mentioned in "Ⅲ．入出力定義" (機能定義書 sheet). Separately collect every
table ID/alias used in 画面設計書's "参照ｴﾝﾃｨﾃｨ" blocks (or, for a batch program with no
画面設計書, the equivalent inline 参照ｴﾝﾃｨﾃｨ blocks inside 機能定義書's own "Ⅳ．機能処理概要", or a
帳票設計書's own get-item blocks) and in every 更新条件表(<TableID>) sheet's header. **Also read
every 更新条件表 sheet's body**, not just its header — the "取得内容"/"取得条件" column (the source
of each column being set) routinely cites *other* tables that feed values into the one being
updated, and it does this **by the source table's Japanese name only, never by table ID** (e.g.
"④ﾜｰｸﾌﾛｰ承認ﾃﾞｰﾀ(No1)" / "ﾜｰｸﾌﾛｰ承認ﾃﾞｰﾀ.会社ｺｰﾄﾞ" — no "TSJAM722" string appears anywhere in that
sheet). A plain text search for a table ID string will silently miss this kind of reference. Before
concluding a table in the I/O list is "declared but unused," don't stop at an ID-string search —
also search every 更新条件表/画面設計書/帳票設計書 sheet for that table's Japanese name (from
Ⅲ．入出力定義's own 名称 column), since that's frequently the only place it's actually written.
This happened for real: `TSJAM722`(ﾜｰｸﾌﾛｰ承認ﾃﾞｰﾀ) was missed this way — an ID-only search across
the whole workbook came back empty, but the table was genuinely referenced (and should carry an R
flag) once searched by name.

**Also follow every "※<共通設計書名>.<項目> 参照" delegation line to its actual target sheet** —
画面設計書/機能定義書 routinely delegate a whole get-item block to a shared common-design workbook
(most often `01_Doc/04_共通設計/07.共通項目取得.xlsx`, e.g. "※共通項目取得.グループ名 参照")
instead of writing the 参照ｴﾝﾃｨﾃｨ inline. The table(s) actually read live inside that delegated
block still count as this program's usage evidence — don't stop at "this program just says
'see common doc'" and treat the referenced table as unused. Open the target workbook, find the
matching named section (search both the WG-specific sheet, e.g. `共通項目取得(工程管理)`, and the
shared `共通項目取得` sheet — items get migrated from the WG-specific sheet to the shared one over
time, noted in a revision comment like "改訂履歴No94" when it happens, so the current live copy
may not be where you'd first expect), and pull its 参照ｴﾝﾃｨﾃｨ table(s) as if they appeared inline
in this program's own doc. This happened for real: `TSJAM726`(ﾜｰｸﾌﾛｰ承認ｸﾞﾙｰﾌﾟ) was missed because
the design doc only wrote "※共通項目取得.ｸﾞﾙｰﾌﾟ名 参照" with no table ID or name anywhere in the
program's own workbook — the actual `TSJAM726` reference only existed inside
`07.共通項目取得.xlsx`'s "(37)ｸﾞﾙｰﾌﾟ名" section, which nothing in the program's own workbook
would surface without deliberately following the delegation. Likewise, a batch program's argument
set may not obviously reach a delegated block's own downstream branches (e.g. a delegated
"移動ﾛｯﾄ(最新)" block that itself references a table via an alias fed by a *different* delegated
block) — if the program-level docs alone can't settle whether that path is actually taken, report
it as a judgment call for the designer rather than asserting it either way.

**Before counting any 参照ｴﾝﾃｨﾃｨ block as evidence a table is used, run the struck-through/grayed-out
formatting scan on that block's table-name column** (see the shared dump doc's section on this) — a
real review missed six false "used but not declared" findings this way on a single large
画面設計書, because a red-and-struck-through `処理区分="..."の場合` branch (and everything under it)
was cited as live evidence without checking formatting first. This matters most on exactly the large
sheets where skipping the scan feels tempting — but don't assume a small block is safe to skip
either: this same mistake recurred on `SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx`, sheet
`帳票設計書(RXJC040)`, on a plain ~10-row block (rows 46-55, "(3)WF情報取得", 参照ｴﾝﾃｨﾃｨ citing
`TXJCM137WF`) that was entirely red+struck through — a review flagged "`TXJCM137WF` used in
帳票設計書 but absent from 機能定義書's CRUD table" without checking the block's formatting first,
when the correct read was "this whole block is deprecated, so there is no live reference to flag as
undeclared at all." The size of the sheet/block is not a reliable signal for whether this check
matters — run it unconditionally, including on small, easily-overlooked blocks, and including
specifically for the "used but not declared in the I/O table" direction of the diff below (it's easy
to only remember this check when confirming a table "is used" at all, and forget it applies just as
much to deciding whether that usage is live enough to demand a CRUD-table entry).

**When comparing table IDs between the I/O table and actual usage, always require an exact
string match — never treat one ID as a match for another just because one is a prefix/substring
of the other.** This project has numeric-suffixed view-ID variants (e.g. `VXJCM004` vs
`VXJCM004_31` vs `VXJCM004_31_ALL`) where the *unsuffixed* base ID commonly has no design file or
real existence of its own at all — only a specific suffixed variant is actually implemented
(mirrors the already-known `WF`-suffix case, e.g. `TXJAM061` vs `TXJAM061WF`, being separate
workbooks). This happened for real: `PXJCO152_処置指示発行.xlsx`'s 機能定義書「Ⅲ．入出力定義」
row 50 (No.12) declares the table ID as bare `VXJCM004`, but every other place this view is
actually used in the same workbook — `更新条件表(TXJCM006)`'s header `[8,68]` and
`ﾁｪｯｸ処理設計書(GXJC152A)`'s 参照ｴﾝﾃｨﾃｨ blocks at `[224,14]`/`[648,14]` — correctly cites the full
`VXJCM004_31`. A substring-based match would call this "consistent" (bare `VXJCM004` is literally
contained in `VXJCM004_31`) and miss that the I/O table's own ID is truncated/incomplete; there is
no design file for a bare `VXJCM004` anywhere in the project, only `VXJCM004_31`-suffixed ones.
Flag any case where an I/O table's declared ID is only a prefix of the ID actually used elsewhere
in the workbook (or vice versa) as its own defect — "declared ID is missing a suffix present in
every live usage" — rather than silently accepting it as the same table. This same exact-match
rule applies to locating the DB layout file below: search for the *exact* full ID (suffix
included), not a fuzzy "contains this substring" search, since a fuzzy search over
`VXJCM004_31_...xlsx`/`VXJCM004_31_ALL_...xlsx` filenames will "find" a file and wrongly appear to
confirm the bare, unsuffixed ID as valid.

Diff the two sets:
- Table used in 画面設計書/更新条件表/帳票設計書 but **absent** from Ⅲ．入出力定義 → flag ("used
  but not declared in the I/O table").
- Table listed in Ⅲ．入出力定義 with a C/R/U/D flag but **never actually referenced** anywhere
  in 画面設計書/更新条件表/帳票設計書 → flag ("declared but unused").
- **For every table you just confirmed IS referenced somewhere (i.e. it passed the previous
  bullet), don't stop there — separately verify each flagged letter (C/R/U/D) against what kind
  of operation the live references actually perform**, table by table. This is easy to skip once
  you've satisfied yourself a table "is used," but "is used" and "is used the way the flags say"
  are different questions — a table can have a live SELECT-style 参照ｴﾝﾃｨﾃｨ block (a real R) while
  its I/O row only carries a C flag from a 更新条件表 sheet, and that gap won't surface unless you
  check flag-by-flag. A real review missed exactly this: `TXJAM025`/`TXJAM026` each had a live,
  unstruck 参照ｴﾝﾃｨﾃｨ reference fetching columns from them (a genuine R), but their I/O rows only
  had a C flag (from their 更新条件表 sheets) — the review confirmed "used, not unused" and moved
  on without checking that the R itself was missing from the flags. Flag any letter present in
  actual usage but absent from the row, and vice versa (a flagged letter with no matching
  operation anywhere).
- Also check whether each table ID appears in
  `01_Doc/06_システム設計書（一覧、管理台帳）/06-06_DB一覧_共通.xlsx` (the DB一覧) and whether its
  ﾃｰﾌﾞﾙﾚｲｱｳﾄ workbook exists. **The DB design folder for a WG is a TOP-LEVEL project folder named
  `<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)` (e.g. `11_工程管理WG\07_データベース・
  ファイル設計書(仮)`) — a sibling of `01_Doc`, NOT nested inside it**, even though most other
  design-doc types live under `01_Doc\...`. A real review once searched under the wrong,
  non-existent path `01_Doc\07_データベース・ファイル設計書(仮)` and wrongly reported 5 real
  tables (TSJAM036, TSJAM037, TSJAM081, TSJAM999, VSJCM137) as having no layout file at all — they
  all existed under the correct top-level WG folder (one in a `PH3` subfolder). Search recursively
  under the correct top-level folder (including any `PH2`/`PH3` subfolders, preferring PH3 > PH2 >
  top-level if a table appears in more than one) before ever reporting a table's layout as missing.
  **基準情報(JA)-prefix common/shared tables have instead been found under
  `01_Doc\07_データベース・ファイル設計書\`** (a different location than the WG-top-level pattern)
  — check there too before concluding a table's layout is genuinely absent. If it's genuinely not
  found anywhere, flag it (note: a table living in a *different* WG's folder is not itself a
  defect, just note it so the user can confirm ownership).

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on.

For each real finding: state the defect in one or two sentences, cite the exact sheet name and cell
coordinates (and the master file it was checked against, e.g. 06-06_DB一覧, a specific ﾃｰﾌﾞﾙﾚｲｱｳﾄ
workbook path) so it's actionable, and give a concrete suggested fix when one is obvious (e.g. "add a
row for table Z with an R flag", "add the missing R flag to TXJAM026's existing row"). Group findings
under a heading only when several findings share one; don't force every finding into a rigid
structure if it doesn't fit cleanly.

Order by impact, most important first — a genuinely missing/undeclared table or a flag that doesn't
match actual usage belongs well before a minor naming nit. When a finding depends on a judgment call
rather than a clear rule (a cross-WG reference you couldn't fully verify, a delegation path you
couldn't confirm the program actually takes, a pattern that could be deliberate rather than an
omission), say so explicitly and mark it as needing the designer's confirmation — don't present it
with the same confidence as a confirmed defect.

Omit entirely, unless the user explicitly asked for a full pass-by-pass audit:
- "確認済み・問題なし" lines for tables that checked out clean
- tallies of how many tables were checked
- narration of your own process (which files you dumped, how you sampled, which scratch path you
  used)

The one exception worth keeping brief: if a whole section couldn't be verified at all (a DB design
folder was missing, a table's layout file couldn't be found anywhere), a one-line note is fine —
that's a real gap in review coverage the designer should know about, not a process detail.
