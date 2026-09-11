---
name: report-design-check
description: Review a program's 帳票設計書(<帳票ID>) sheets — the report/print specification that no other REV skill looks at in its own right. Checks that the report is registered in 06-03_帳票一覧_<WG>.xlsm under this program, that Ⅱ．帳票仕様's fixed spec rows (用紙ｻｲｽﾞ/明細部/改ﾍﾟｰｼﾞ条件/0件出力/出力順/抑止項目…) are actually filled in rather than left blank, that every 参照先 alias used in Ⅲ．編集仕様 is defined by a 参照ｴﾝﾃｨﾃｨ in Ⅰ．帳票出力条件, that every printed value traces back to a 取得項目 (and no fetched item goes unprinted), that 画面項目ID values on print items are registered in 82.画面項目辞書, and that the ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ(<帳票ID>) layout sheet agrees with the print-item list. Distinct from design-doc-io-table-check (which only asks whether a 帳票設計書's tables reach the CRUD list) and xlsx-db-column-check (column existence). Use when the user asks to REV 帳票設計書/帳票仕様/印字項目. When the user asks to REV a single design-doc workbook without naming which checks they want, the entry point is `rev-program-review`: it first asks the user, checkbox-style, which of the 8 single-program checks to run, then runs only those as one combined pass. Do not launch all eight yourself. Run this skill standalone only when it was one of the selected checks, or when the user asked for this check by name.
---

# report-design-check

Reviews every `帳票設計書(<帳票ID>)` sheet in one program's design-doc workbook. Reports are a real
blind spot in the existing check set: `design-doc-io-table-check` reads 帳票設計書 sheets only to
harvest table IDs for the CRUD list, `xlsx-db-column-check` only to harvest column names, and
`naming-standard-compliance` only to check the 帳票ID's shape. Nothing checks whether the report
specification itself is complete.

It is also where the findings are. The project's 第3フェーズ基本設計 指摘傾向レポート puts 帳票・
ファイル出力 at 26% of one designer's findings and 18% of another's — and those two designers own the
report-heavy programs.

**Called from `rev-program-review`?** Then the workbook is already dumped and the scope is already
chosen — don't re-dump it, don't unconditionally launch the other checks, and don't post a
per-check status update. This skill runs on its own when it was one of the selected checks, or when
the user asked for this check by name.

**A program with no `帳票設計書(*)` sheet is not in scope.** Say so in one line and stop — do not
substitute the `ﾌｧｲﾙ出力仕様書(F*)` sheet, which is a different document with a different shape.

## Environment constraints (important)

**Where the shared docs live:** the `_shared/*.md` files ship **inside this plugin**, in the
`_shared/` folder next to this skill's own directory (`<plugin root>/skills/_shared/`) — **not** in
`~/.claude/skills/_shared/`, which does not exist on a normal install. Resolve every `_shared/...`
reference below against that folder; if it doesn't resolve, glob
`**/rev-skills/**/skills/_shared/<filename>` and read the hit.

Read `_shared/xlsx-excel-com-dump.md` first (the Excel COM dump — this machine has no working
Python/Node), and `_shared/reference-index.md` for the 画面項目辞書 index used by check C5.

## Sheet anatomy (measured, not assumed)

Verified against the five `帳票設計書(*)` sheets of
`01_Doc\08_機能定義書\11_工程管理\PHASE3\SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx`. Anchor on the **labels**;
the column numbers below are the common case, not a contract.

```
[4,5]=帳票ID   [4,10]=RXJC041  [4,15]=添付制限情報      ← the sheet's identity
[6,2]=Ⅰ．帳票出力条件
  [14,4]=参照ｴﾝﾃｨﾃｨ [14,12]=A  [14,14]=TSJAM081:帳票ﾏｽﾀ  ← alias → table
                    [15,12]=ZY [15,14]=引数(呼出し元画面)ﾛｸﾞｲﾝ情報
  [17,4]=取得項目   [18,6]=項目名 [18,17]=取得内容
  [21,4]=検索条件   [22,6]=A.会社ｺｰﾄﾞ [22,24]== [22,28]=ZY.会社ｺｰﾄﾞ
  [25,4]=ｿｰﾄ順
  [27,4]=取得件数 [27,9]=1件 [27,15]=取得できない場合 [27,24]=-
[45,2]=Ⅱ．帳票仕様        ← a fixed row-label list, see C2
[63,2]=Ⅲ．編集仕様
  [65,3]=1.処理の流れ      ← numbered narrative
  [84,3]=2.ﾍｯﾀﾞｰ情報 / [110,3]=3.明細情報 …   ← one section per output area
    [92,4]=(1)段落ﾀｲﾄﾙ / (2)ﾍｯﾀﾞｰ / (3)明細
      [94,5]=No. [94,7]=出力項目名 [94,17]=出力内容
      [95,17]=参照先 [95,23]=項目名・出力値 [95,45]=編集方法 [95,54]=画面項目ID
      [96,5]=1 [96,7]=品目ｺｰﾄﾞ … [96,17]=- [96,23]="品目ｺｰﾄﾞ" [96,45]=文字列 [96,54]=SJC0794\nXJC8049
```

Parsing rules that matter:

- **The print-item table has a two-row header**: `No./出力項目名/出力内容` on the first, and
  `参照先/項目名・出力値/編集方法/画面項目ID` underneath, splitting 出力内容. Anchor on the second row
  for the value columns.
- **`出力項目名` spans merged cells** and therefore repeats across ~10 columns (7 through 16) in the
  dump. That is a merge artifact, not ten values — take the first and ignore the rest. Do not report
  it as a formatting irregularity either; that is
  `design-doc-formatting-consistency`'s call to make, and this shape is the norm here.
- **A `画面項目ID` cell often holds two IDs separated by a newline** (`SJC0794\nXJC8049` — the label
  item and the value item). Split on newline and check each.
- **The ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ sheets are separate sheets**, named `ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ(<帳票ID>)` and sometimes with
  a variant suffix (`ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ(RSJC033)CP1,CMOS`, `…(RSJC033)識別CP1,CMOS`). They are wide
  (~211 columns) grid mock-ups of the printed page.
- **`bk_`-prefixed and hidden sheets are out of scope** (`bk_帳票設計書(RSJC034)` etc.), per the
  shared dump doc's default. A `bk_` sheet for a report with no live sheet is not a live report.
- Struck-through content is already resolved out of the dump; don't pull it back from
  `_DELETED_DIGEST.txt`.

## Procedure

### 1. Get the workbook dump, and enumerate the reports

From the shared dump (or dump per `_shared/xlsx-excel-com-dump.md`), list every live
`帳票設計書(<帳票ID>)` sheet and the `ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ(<帳票ID>)` sheets that pair with them.

### 2. Read the 帳票一覧 registry for the report IDs' WG

Route by the **report ID's own JOBコード**, exactly as `_shared/reference-index.md` routes
screen-item IDs: `R` + `XJC`/`SJC` → `01_Doc\06_システム設計書（一覧、管理台帳）\06-03_帳票一覧_工程管理.xlsm`,
`XJA`/`SJA` → `_基準情報`, `XJB`/`SJB` → `(XJB_受注出荷)`/`(SJB_受注出荷)`, `XJD`/`SJD` → `_品質管理`,
`XJZ`/`SJZ` → `_共通`. A program's own reports can span two prefixes and still live in one file —
`06-03_帳票一覧_工程管理.xlsm` holds both `RXJC*` and `RSJC*` — so route per ID and deduplicate the
files, rather than assuming one file per program.

Dump it through the cross-session cache (`_shared/xlsx-excel-com-dump.md`). The registry sheet is
`帳票一覧(<WG>)`: header at row 4, data from row 5 —
`No.(1) | 帳票ID(2) | 帳票名称(3) | 編成(4) | RL(5) | BF(6) | BL(7) | 区分(8) | 備考(9) | 計画書No.(10)`.
The `備考` column carries the **owning program's name** (`処置指示発行(ｻﾌﾞﾌﾟﾛ)`), which is what C1
cross-checks. Trailing rows can carry an ID with an empty name (`RSJC037`, `RSJC038`) — reserved
slots, not registrations.

**A partially-struck name cell is a rename-in-place, not a dead row.** Apply
`_shared/reference-index.md`'s `LiveText` cascade. This exact registry is where the mistake was made
for real: `RSJC033` and `RSJC035` were reported as 帳票一覧未登録 by three separate agents when their
rows existed and had merely been renamed (`前工程処置指示書（CP1,CMOS）` → `処置指示書(前工程)`). Drop a
row only when the **帳票ID** cell has no live text left.

### 3. Run the checks

**C1 — 帳票一覧への登録.** Every live 帳票ID in the workbook must have a live registry row, and that
row's `備考` should name this program. Report both directions: a report designed here but not
registered, and a registry row attributed to this program with no 帳票設計書 sheet (usually a report
that was dropped from the design without being withdrawn from the registry — check whether the sheet
exists as `bk_` before calling it missing). Also compare 帳票名称: registry vs the sheet's `[4,15]`.

**C2 — Ⅱ．帳票仕様の記入漏れ.** This section is a fixed label list, and a blank or `-` on the wrong
row is the single most common report finding. Walk the labels present in the sheet — the observed
full set is 出力ｺｰﾄﾞ / 一時ﾌｧｲﾙ名(共通) / ﾀﾞｳﾝﾛｰﾄﾞﾌｧｲﾙ名(PG個別) / 出力方法 / 用紙ｻｲｽﾞ / 明細部 /
合計部 / 改ﾍﾟｰｼﾞ条件 / ﾍﾟｰｼﾞﾘｾｯﾄ条件 / ﾍﾟｰｼﾞｶｳﾝﾄ方法 / 抑止項目 / 0件出力 / ﾌｫﾝﾄ / 出力順 / 備考 —
and check:

- `出力ｺｰﾄﾞ` must equal the sheet's own 帳票ID.
- `用紙ｻｲｽﾞ` and `出力方法` must be concrete. Empty is a finding.
- **`0件出力`** — `なし` is a decision, empty is an omission. If the report has a 明細 section fed
  by a multi-row fetch, an empty 0件出力 means nobody decided what prints when the query returns
  nothing. This is the row most often left blank.
- **`改ﾍﾟｰｼﾞ条件` / `明細部` / `出力順`** — a report with a 明細 section and `明細部 = -` or an empty
  出力順 cannot be implemented deterministically: the coder has no row count per page and no sort.
  Cross-check 出力順 against the `ｿｰﾄ順` of the Ⅰ．帳票出力条件 block that feeds the detail; `ｿｰﾄ順 =
  なし` together with a multi-row 明細 is itself a finding (non-deterministic print order).
- A `-` on a row where the report plainly needs a value (合計部 `-` on a report whose 明細 has a
  総計 column) is worth a line as a lower-confidence judgment call.

Do **not** flag `-` on 明細部/合計部/出力順 for a single-page, header-only report — read the Ⅲ．編集仕様
sections first and only raise these where a 明細 section actually exists.

**C3 — 参照先エイリアスの定義.** Collect every alias used in Ⅲ．編集仕様's `参照先` column and in the
`項目名・出力値` cells' `<alias>.<column>` references, plus the `検索条件` right-hand sides. Each must
be defined as a `参照ｴﾝﾃｨﾃｨ` in the Ⅰ．帳票出力条件 block that precedes it (`A` → `TSJAM081:帳票ﾏｽﾀ`,
`Y`/`ZY` → 引数). **Aliases are block-scoped and get reused** — `A` means a different table in
`(2)帳票ﾃﾝﾌﾟﾚｰﾄ取得` than in `(3)工程情報取得`. Resolve each reference against its own block before
calling an alias undefined, and when the 編集仕様 references an alias across blocks without saying
which, that ambiguity is the finding.

**C4 — 取得項目と印字項目の双方向照合.**
- Every `<alias>.<column>` printed in Ⅲ must appear in that block's `取得項目` list. A printed value
  that was never fetched is a genuine implementation hole.
- Every `取得項目` entry should be printed somewhere in Ⅲ (or consumed by another block's 検索条件).
  An unprinted fetch is usually a leftover from an earlier revision — lower confidence, report it
  as such.
- A `項目名・出力値` that is a bare quoted literal (`"品目ｺｰﾄﾞ"`) is a fixed label, not a data
  reference. Don't chase it into the 取得項目 list.

**C5 — 画面項目IDの登録.** Every ID in a `画面項目ID` cell must be registered in the
`82.画面項目辞書_*.xlsx` its prefix routes to — build/read the index per `_shared/reference-index.md`
(do **not** dump the dictionary workbook wholesale). Split multi-ID cells on newline. Also flag the
opposite shape: a print item in a 段落ﾀｲﾄﾙ or ﾍｯﾀﾞｰ table with **no** 画面項目ID at all where its
siblings in the same table all have one — the column exists so that the printed label can be
multilingual, and a blank means the label is hardcoded.

**C6 — ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄとの整合.** For each report with a `ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ(<帳票ID>)` sheet, check that
every 出力項目名 in the print-item tables appears somewhere on the chart, and that the chart has no
labelled field absent from the print-item tables. These sheets are wide grids, so compare on item
**names**, not positions, and keep this check's confidence honest — a chart cell can legitimately
carry a caption that is not a print item. Report a missing item as "帳票設計書にあるがﾚｲｱｳﾄにない"
(and the reverse) rather than asserting a defect. A report whose chart sheet is missing entirely,
while its siblings all have one, is worth a line on its own.

**C7 — 処理の流れの参照先.** Ⅲ．編集仕様's `1.処理の流れ` cites blocks by number (`①(2).ﾌｧｲﾙﾃﾞｰﾀ①を
取得する`) and cites the 機能定義書 (`※機能定義書(SXJCB147).Ⅳ．機能処理概要.A-②を参照`). Check those
pointers resolve: the `(2)` block exists in Ⅰ．帳票出力条件, and the cited 機能処理概要 step exists in
the 機能定義書 sheet. A dangling pointer here is cheap to find and expensive to hit in coding.

### 4. Verify before reporting

Re-read the cells behind each candidate finding. The traps specific to this sheet shape:

- treating the merged `出力項目名` repetition as ten separate items;
- resolving a block-scoped alias against the wrong block (C3) — the most likely source of a wrong
  "undefined alias" or "unfetched column" finding;
- calling a renamed-in-place 帳票一覧 row unregistered (see step 2 — this has happened).

## Reporting

Write in the user's language, for the designer who owns the doc.

Group by 帳票ID. For each finding: name the report, cite the sheet and cell
(`帳票設計書(RXJC041) [57,18]`) with a clickable `[filename](relative/path)` link, and state the
consequence in one clause (0件時の出力が未決定 / 印字項目が取得されていない / 帳票一覧に未登録).
Order by severity: C1 and C4's unfetched-print items first, then C2's blank spec rows, then the
rest. Mark C6 and the unprinted-fetch half of C4 as lower-confidence judgment calls.

Omit: reports that checked out clean, a tally of how many reports/items were compared, and
narration of which files you dumped. One process fact is worth a line — a 帳票一覧 file that
couldn't be located for a report's prefix, since that leaves C1 unchecked for those reports.
