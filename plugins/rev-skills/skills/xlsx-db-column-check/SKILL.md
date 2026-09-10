---
name: xlsx-db-column-check
description: Check whether a function-definition Excel doc (機能定義書/画面設計書/更新条件表) references DB columns that don't actually exist in the corresponding table-layout (テーブルレイアウト) Excel files. Distinct from design-doc-io-table-check, which checks whether a table is *declared* in the Ⅲ．入出力定義 CRUD list at all — this skill instead checks whether the *columns* referenced within an already-used table actually exist. Also flags a project-specific anti-pattern in 検索条件保存マスタ-style generic tables (e.g. TXJAM100): persisting both a master-entity code (品目コード等) AND its master-derived display name (KC品名等) together, when only the code should be stored and the name should come from a JOIN at read-time — but ONLY when the screen shows that name as a Label; a name the user types into a TextBox is an independent search condition and persisting it is correct. Use when the user asks to verify a program's design doc against DB/file design docs, e.g. "このファイルが使っているカラムが、DB設計書のファイルに存在するか確認して" or "存在しないカラムを使っていたら教えて". When the user asks to REV a single design-doc workbook without naming which checks they want, the entry point is `rev-program-review`: it first asks the user, checkbox-style, which of the 6 single-program checks to run, then runs only those as one combined pass. Do not launch all six yourself. Run this skill standalone only when it was one of the selected checks, or when the user asked for this check by name.
---

# xlsx-db-column-check

Cross-checks the DB columns referenced by a function/screen design Excel workbook against the
actual column list defined in the project's table-layout ("テーブルレイアウト") Excel files — see
the frontmatter `description` above for what counts as a miss and how this differs from
`design-doc-io-table-check`. Developed against the common Excel template shared across WGs in this
codebase (機能定義書 / 画面設計書 / 更新条件表 / ﾃｰﾌﾞﾙﾚｲｱｳﾄ sheets); generalizes to any
`<機能定義書xlsx>` + `<DB設計書folder>` pair, not just one specific program.

**Scope selection comes first.** When the user asks to REV a single program's design-doc workbook
without naming specific checks, `rev-program-review` is the entry point: it presents the 6
single-program checks as a checkbox list (`AskUserQuestion`, multiSelect), then runs only the
selected ones as one combined pass, dumping the workbook once up front and sharing the text with
every check (see `_shared/xlsx-excel-com-dump.md`'s "dump once, share the text" section). Do **not**
unconditionally launch all 6 yourself, and do not post a per-check status update — the combined
report is posted once, after every selected check has finished. This skill runs on its own when it
was one of the selected checks, or when the user asked for this check by name.

## Environment constraints (important)

**Where the shared docs live:** the `_shared/*.md` files ship **inside this plugin**, in the
`_shared/` folder next to this skill's own directory (`<plugin root>/skills/_shared/`) — **not** in
`~/.claude/skills/_shared/`, which does not exist on a normal install. Resolve every `_shared/...`
reference below against that folder; if it doesn't resolve, glob
`**/rev-skills/**/skills/_shared/<filename>` and read the hit. Do not skip it and improvise the
Excel COM dump instead: that doc carries the guard against attaching to — and then `Quit()`-ing —
the user's own live Excel session, the strikethrough-exclusion scan, and the reference-file cache.

Read `_shared/xlsx-excel-com-dump.md` first — it has the PowerShell + Excel COM
script this skill (and its sibling review skills) use to read `.xlsx` files, since this machine has
no working Python/Node to use a library like openpyxl.

## Procedure

### 1. Dump the target design workbook to text

Follow `_shared/xlsx-excel-com-dump.md` to dump every worksheet of the target workbook to
plain-text files under the session scratchpad directory, one file per sheet.

### 2. Extract the referenced tables and columns from the design doc

Read the dumped sheets with the Read tool (not Bash cat) and look for these standard sections:

- **機能定義書 sheet, "Ⅲ．入出力定義"**: a table listing every DB table/file the program touches
  (ID, 名称, C/R/U/D, 用途). This is the master list of tables in scope — use it to know what to
  look up in step 3, and to note any table whose design file can't be found in the target folder.
- **画面設計書 sheet, "Ⅲ．画面表示仕様"**: numbered subsections (e.g. "(3)加工手順",
  "(4)出力実績表項目取得") each with a "参照ｴﾝﾃｨﾃｨ" block assigning aliases (A/B/C/... plus Y=画面,
  Z=ﾛｸﾞｲﾝ情報) to tables, followed by 取得項目 (columns fetched), 検索条件 (search conditions),
  結合条件 (join conditions), and ｿｰﾄ順 (sort columns) — each line references `<alias>.<column
  name>`. Collect every `<column name>` per real table alias (ignore Y/Z, those are screen/session
  fields, not DB tables).
- **更新条件表(<テーブルID>) sheet(s)**: a numbered list of columns for one target table being
  inserted/updated, with a "取得内容" column showing where each value comes from. Collect the full
  項目名 list for that table.

Build a per-table list: `{ table_id: [column names referenced] }`.

**Struck-through/grayed-out rows are already excluded — do NOT run a formatting scan of your own.**
The dump script resolves this while the workbook is open (see
`_shared/xlsx-excel-com-dump.md`'s "Excluding struck-through / grayed-out rows from review"), so a
deprecated 参照ｴﾝﾃｨﾃｨ block and every column reference under it (取得項目/検索条件/結合条件/ｿｰﾄ順)
are simply absent from the `.txt`. Report a brief count of excluded cells — the dump prints
`dead=`/`partial=` per sheet — rather than listing them.

**The one thing you must still be deliberate about: a partially-struck cell arrives already reduced
to its live text, and that live text is the column name to match.** Never reconstruct the raw string
from `_DELETED_DIGEST.txt` and match on that. This is the single most repeated false finding in this
skill's history — `注意事項備考` reported as "not a column in TXJCM137" when only `注意事項` was live
(`SXJCB147`, `帳票設計書(RSJC035)`), `COUNT(A.層数層No)` "fixed" to `COUNT(A.層No)` when `層No` was
already the live text (`PXJCO124`), and `③④` treated as two entity references when only `④` was live
(`PSJCO205`). All three came from matching raw text; none can occur when matching what the live dump
gives you.

### 3. Locate and dump the actual DB design files

**Quick reference — resolving a table ID to its one authoritative layout file:**

1. Search root: `<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)` — a TOP-LEVEL project
   folder, sibling of `01_Doc`, **NOT** nested inside it. Search recursively (it may have `PH2`/
   `PH3`/`ファイルレイアウト` subfolders).
2. Match the design doc's ID **exactly** — never a prefix/substring match. A `WF` suffix
   (`TXJAM061` vs `TXJAM061WF`) or a numeric suffix (`VXJCM004` vs `VXJCM004_31`/`VXJCM004_31_ALL`)
   makes it a different table; the bare/unsuffixed form often has no design file of its own at all.
   **A filename glob is not an exact match** — `<ID>_*.xlsx` also matches suffixed *other* tables,
   because `_`-suffixed IDs are themselves real. Confirm the ID written in **the `ﾃｰﾌﾞﾙﾚｲｱｳﾄ` sheet's**
   `A6` cell (the one under the `ﾃｰﾌﾞﾙID` label in `A5`) and reject any candidate that disagrees,
   whatever its phase folder. **`A6` is sheet-dependent — read it from the wrong sheet and you reject
   the right file.** A layout workbook also carries `JAGﾃｰﾌﾞﾙﾚｲｱｳﾄ` and one or more `旧ﾃｰﾌﾞﾙﾚｲｱｳﾄ`
   sheets; in `TXJCM003_製造ｵｰﾀﾞｰ.xlsx` those hold `TXJCM003`, `FDMBM03`, `FDCJM03` and `No.`
   respectively. Resolving `SXJCB147`'s tables
   by filename alone picked the wrong file for three of them —
   `TXJCM003_B_ｵｰﾀﾞｰ投入出荷予定.xlsx` for `TXJCM003` (製造ｵｰﾀﾞｰ),
   `TXJCM007_B_移動ﾛｯﾄ構成取消履歴.xlsx` for `TXJCM007` (移動ﾛｯﾄ構成),
   `TXJAM008_B_共通ｺｰﾄﾞﾏｽﾀ(選択肢).xlsx` for `TXJAM008` (共通ｺｰﾄﾞﾏｽﾀ) — each of which would have
   produced false "column does not exist" findings against another table's column list.
   `TXJAM008_IN_共通ｺｰﾄﾞﾏｽﾀ受信.xlsx` (`TXJAM008_IN`) is a fourth trap on that same table. The `A6`
   check overrides the phase order in step 3: a higher-priority file whose `A6` disagrees is not the
   file.
3. If the same ID exists under more than one phase folder, resolve by **PH3 > PH2 > top-level** —
   never by file-modified date.
4. Also check the flat `01_Doc\07_データベース・ファイル設計書\<table>.xlsx` (no `(仮)`) for an
   independent copy — this is a real, separate location, distinct from (and not to be confused
   with) the wrong, non-existent `01_Doc\07_データベース・ファイル設計書(仮)` path.
5. For a table owned by **11_工程管理WG**, trust its own WG-folder copy — no further
   cross-checking needed. For a table owned by **any other WG**, also check the flat `01_Doc\...`
   copy from step 4; if the two disagree, prefer the one with a strict column superset (treat a
   non-superset disagreement — renamed/reordered columns — as a discrepancy to report, not resolve
   yourself).
6. Never search `<WG番号>_<WG名>WG\開発DDL作成用<date>\` for a table's layout — always excluded
   from a REV per the shared dump doc's "Folders to always exclude" list, regardless of which WG
   owns the table.
7. If a table's file is genuinely not found anywhere under the specified folder, that's *not* a
   "missing column" finding — report it separately as "design file not present" (usually means the
   table is common/shared infrastructure living in another WG's or the common docs' folder).

**Background — why each rule exists** (skip on a routine run; kept for onboarding/reference):

- **Wrong search root.** A real run once constructed the path as
  `01_Doc\07_データベース・ファイル設計書(仮)` (mirroring the `01_Doc`-nested pattern of other doc
  types), got zero matches, and wrongly reported 5 real tables (TSJAM036, TSJAM037, TSJAM081,
  TSJAM999, VSJCM137) as having no layout file at all — all 5 existed under the correct top-level
  WG folder (one in a `PH3` subfolder), and every referenced column in them checked out fine.
- **Substring matching instead of exact-ID matching.** This project has views like
  `VXJCM004_31`/`VXJCM004_31_ALL` where the bare, unsuffixed ID has no design file of its own.
  `PXJCO152_処置指示発行.xlsx`'s I/O table cites bare `VXJCM004`, while every live usage elsewhere in
  the same workbook correctly uses `VXJCM004_31` — a substring search for `*VXJCM004*` matches the
  `VXJCM004_31...` filenames and wrongly looks like the bare ID resolved to a real file. If the doc
  says `VXJCM004` and only `VXJCM004_31` exists on disk, that mismatch is itself a finding (the
  doc's own ID is incomplete/wrong), not a confirmed match.
- **Guessing phase precedence from file dates instead of PH3 > PH2 > top-level.** An earlier,
  less reliable attempt compared file-modified timestamps and produced a real false-positive finding
  when a `PH2` copy of `TXJCM058` turned out to be the current, more complete one despite the
  top-level copy looking no different at a glance. Do not feed more than one copy of the same table
  ID into the comparison in step 4 — resolve to exactly one file per table ID before dumping.
- **The `開発DDL作成用` exclusion (rule 6) is a standing user decision, not a discovery of "no
  value there."** It was originally added after a real case where `TSJBM012`(梱包Noﾏｽﾀ)/
  `TSJBM013`(梱包紐付けﾏｽﾀ), both owned by 受注出荷WG, existed in three places, and the owning-WG copy
  under `05_受注出荷WG\20_基本設計\07_データベース・ファイル設計書(仮)\` was missing a `梱包Rev`/
  `KONPO_REV` column that a `11_工程管理WG\開発DDL作成用0930\` snapshot copy did have. If a similar
  gap resurfaces under the current exclusion, note it as a coverage limitation in the report rather
  than searching the excluded folder anyway.

Once every table's ID has resolved to exactly one layout file (step 3's PH3 > PH2 > top-level
precedence, applied per table), **dump all of them in one pass using the "Batch variant" script in
`_shared/xlsx-excel-com-dump.md`'s cross-session cache section**, each entry with
`OnlySheetPatterns = @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ")` — a table-layout workbook commonly has 3-4 sheets (改訂履歴, the live
ﾃｰﾌﾞﾙﾚｲｱｳﾄ, and one or two old JAG_/旧-prefixed superseded copies), and this is the only one this
check ever reads, so there's no reason to pay the Excel COM/cache cost of the others — and a program
easily has 10+ tables in its I/O list, so batching means Excel launches at most once for this whole
step (often zero times, once the cache is warm from an earlier REV or from `db-design-cross-consistency`
having already dumped the same tables). That sheet has
a fixed layout: row 6 = `[6,1]=<TableID>`/`[6,6]=<Table name>`, header at row 7
(`No. | 項目名 | 項目ID | 属性 | 桁数 | DB桁 | I01... | notnull | 備考`), and one data row per
column starting at row 8. Column `[r,3]` is the 項目名 (Japanese column name) — this is the
authoritative list of columns that actually exist.

### 4. Compare and report

For each `{table_id: [referenced columns]}` from step 2, check each referenced column name against
the actual 項目名 list from step 3's dump for that table.

- **Exact miss** (referenced name has no match at all in the real table, including no
  similarly-named field holding the same kind of data): report as a clear finding — cite the
  design-doc sheet/cell where it's referenced and state that the table has no such column.
- **Likely rename** (referenced name differs only by a prefix/suffix or synonym, e.g. `工程` vs
  `対象工程`): mention as a secondary, lower-confidence note — it may just be documentation
  shorthand rather than a real defect — but still flag it so the user can judge.
- **Suspicious cross-reference**: when a design doc says `<alias>.<column>` and column is genuinely
  absent from that alias's table, but a *different* table in the same subsection has a
  same-purpose column under a different name (e.g. only table A has `実績表項目名` but the doc
  points at table C for it), call this out explicitly — it usually indicates the doc's join/alias
  is wrong, not just a naming variant.

### 5. Redundant code+name persistence in 検索条件保存マスタ-style generic tables

Some common tables in this project (confirmed for `TXJAM100`:検索条件保存ﾏｽﾀ, a shared table that
persists each user's last-used search filters across many different programs' screens) are
generic-column designs — their real ﾃｰﾌﾞﾙﾚｲｱｳﾄ has no per-field schema at all, just a long run of
polymorphic `項目1`...`項目N` columns (150 of them for TXJAM100), each holding whatever a given
program's 更新条件表 decides to put there. **The project's own design rule for this kind of table:
when a search field is a code that references a real master entity (品目ｺｰﾄﾞ, 作業場ｺｰﾄﾞ, 工程ｺｰﾄﾞ,
取引先ｺｰﾄﾞ, etc.), only the code itself should be persisted into the generic table — the
corresponding display name (KC品名, 作業場名, 工程名, 取引先名, etc.) must be re-fetched via a JOIN
against the real master table at read-time, never stored redundantly alongside the code.** A
classification/区分 code whose paired label comes from 共通ｺｰﾄﾞﾏｽﾀ is a different, correctly-handled
case — this project's own convention there is to persist the code alone and annotate it
"(区分値のみ)" in the 取得内容 column, which is *not* a violation (it's the same discipline applied
correctly).

**Detection procedure** (run this specifically whenever a 更新条件表 targets `TXJAM100` or another
table you've confirmed follows the same generic-column/検索条件保存 convention): for every row whose
INSERT/UPDATE column shows a real, non-`-` source (i.e., it's actually being persisted), extract the
plain-text field label from the 取得内容 value (e.g. `G1)作業場名` → `作業場名`; strip the leading
`<画面エリア>)` prefix first). Classify each label as **code-type** (ends in `ｺｰﾄﾞ`/`CD`) or
**name-type** (ends in `名`/`名称`/`略式名`/`品名`, or is a known name field like `KC品名`). For every
code-type label, strip the code suffix to get its entity stem (e.g. `作業場ｺｰﾄﾞ` → `作業場`) and check
whether any name-type label in the same 更新条件表 shares that stem (`作業場` → `作業場名`) — or, for
the `品目`/`KC品名` case specifically, treat `KC品名` as `品目`'s paired name even though the surface
text doesn't share the stem literally (this pairing is specific to this project's terminology: `KC`
prefix names are this project's product-name field for a 品目ｺｰﾄﾞ). A code+name pair that are **both**
actually persisted (neither is `-`) is a *candidate*, not yet a finding — it must still clear the
control-type gate below. A code-type label whose paired name is annotated "(区分値のみ)" or has no
persisted name-type sibling at all is correctly designed — don't flag it.

**MANDATORY control-type gate — a persisted name is only a violation when the screen shows that name
as a Label.** Per explicit user direction, this gate decides the finding and there is no exception to
it. Before reporting any candidate pair, look the **name** item up by 画面項目名 in that screen's
`Ⅴ．画面項目定義` and read its 属性 (control-type) column — in the standard layout it is array column
18, with `画面項目名` at column 5 and 初期値 at column 36; resolve it by the row-465-style header
labels rather than trusting those positions:

- **属性 = `Label`** (or otherwise display-only) → the name is *derived from the code* via a master
  lookup and has no independent existence. Persisting it duplicates master data. **This is the
  finding.**
- **属性 = `TextBox`** (or any editable input: ComboBox a user picks, etc.) → the name is **its own
  independent search condition** that the user types, not a master-derived echo of the code. The
  program must persist it, because it is part of the filter the user entered and there is nothing to
  JOIN it back from — a partial/LIKE name search has no code to re-derive it. **Persisting it is
  correct. Do not flag it.**

The surface text is identical in both cases (`工程ｺｰﾄﾞ` + `工程名` persisted side by side), so a
name-pairing scan alone cannot tell a real defect from correct design. Skipping the gate turns every
screen that offers name-based text search into a page of false findings.

Confirmed for real on `PXJCO128_ﾛｯﾄ停止指示登録.xlsx`, where the gate flips the verdict on all four
candidates and the design turns out to be **entirely correct and internally consistent**:

| 画面項目 (`画面設計書(GXJC128A)` Ⅴ) | 属性 | TXJAM100 に保存? | 判定 |
|---|---|---|---|
| ﾛｯﾄ停止指示者名 `[484,5]` | TextBox | 保存 `[40,36]` | 正 — 独立した検索条件 |
| 工程名 `[487,5]` | TextBox | 保存 `[42,36]` | 正 — 独立した検索条件 |
| 停止者名 `[492,5]` | TextBox | 保存 `[45,36]` | 正 — 独立した検索条件 |
| 解除者名 `[501,5]` | TextBox | 保存 `[50,36]` | 正 — 独立した検索条件 |
| 停止工程GRP名 `[479,5]` | **Label** | 保存せず | 正 — ｺｰﾄﾞのみ保存 |
| 取引先名 `[508,5]` | **Label** | 保存せず | 正 — ※1復元表 `[538,17]` が `(2).取引先略式名` から再取得 |

The two Labels are exactly the ones left out of the persisted set, and `取引先名`'s restore rule
points at a delegated master lookup rather than a saved 項目N — the read-time-JOIN discipline this
rule is about, applied correctly. A review that reported the four TextBox names as redundant
persistence (as one did before this gate was written) is producing false findings on a correct
design.

Note this also makes `PSJCO304`'s confirmed violation above stronger evidence rather than weaker:
re-check those pairs' 属性 before citing them, and cite the 属性 in the finding itself so the designer
can see the gate was applied.

Confirmed for real on `PSJCO304_着手ﾒｯｾｰｼﾞﾒﾝﾃﾅﾝｽ.xlsx`, 更新条件表(TXJAM100): items 6/7 both persist
`作業場ｺｰﾄﾞ` *and* `作業場名`, items 10/11 both persist `工程ｺｰﾄﾞ` *and* `工程名`, items 12/13 both
persist `取引先ｺｰﾄﾞ` *and* `取引先名`, and items 3/8 (non-adjacent, same 更新条件表) persist both
`品目ｺｰﾄﾞ` and `KC品名` — every one of these is the project's forbidden redundant-name pattern, sitting
right next to two other items (作業場表示区分/工程表示区分) in the very same table that correctly
persist "(区分値のみ)" — proving the rule was known and simply not applied consistently.

### 6. Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on. Write in the user's language.

For each real finding (exact miss, suspicious cross-reference, redundant code+name persistence, or a
likely-rename worth a second look): name the table and column, cite the design-doc sheet/cell where
it's referenced with a clickable `[filename](relative/path)` link, and state plainly what's wrong (no
such column exists / a different table has the matching column under another name / both the code
and its master-derived name are being persisted when only the code should be). Order by confidence —
exact misses, suspicious cross-references, and redundant code+name pairs first, likely-renames after,
clearly labeled as lower-confidence judgment calls rather than confirmed defects.

Omit entirely: tables that checked out clean (don't list "table X: OK"), a running tally of how many
tables/columns were checked, and narration of your own process (which files you dumped, which
folders you searched). The one exception worth a one-line mention: a table whose design file
couldn't be found in the specified folder at all — that's a real coverage gap the designer should
know about, not a process detail.
