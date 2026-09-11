---
name: update-condition-completeness
description: Check a program's 更新条件表(<TableID>) sheets against the table's actual ﾃｰﾌﾞﾙﾚｲｱｳﾄ — every column of the table present in the right order, no notnull column left unset on INSERT, every primary-key column set on INSERT and carried as [KEY] on UPDATE/DELETE, and the house conventions for the 共通項目 block (登録者/登録日時 only at INSERT, 更新者/更新日時 on both, 排他ﾌﾗｸﾞ 1 / +1, 更新ﾌﾟﾛｸﾞﾗﾑID=画面ID) honored. Distinct from its siblings: `xlsx-db-column-check` asks whether a referenced column exists at all, `design-doc-io-table-check` asks whether the table is declared in the Ⅲ．入出力定義 CRUD list — this skill asks whether the UPDATE/INSERT/DELETE specification for an already-declared table is complete enough to code from. Use when the user asks to check 更新条件表 completeness, NOT NULL/主キー/共通項目の設定漏れ, or 登録日時が更新されていないか. When the user asks to REV a single design-doc workbook without naming which checks they want, the entry point is `rev-program-review`: it first asks the user, checkbox-style, which of the 8 single-program checks to run, then runs only those as one combined pass. Do not launch all eight yourself. Run this skill standalone only when it was one of the selected checks, or when the user asked for this check by name.
---

# update-condition-completeness

Reviews every `更新条件表(<TableID>)` sheet in one program's design-doc workbook against the
authoritative `ﾃｰﾌﾞﾙﾚｲｱｳﾄ` of the table it updates. The question this check answers is **"could a
coder implement this INSERT/UPDATE/DELETE from the sheet alone, and would the row it writes satisfy
the table's own constraints?"** — not "does this column exist" (that is `xlsx-db-column-check`) and
not "is this table declared in the CRUD list" (that is `design-doc-io-table-check`).

This is the second-largest finding category in the project's own review history: the
第3フェーズ基本設計 指摘傾向レポート puts 更新条件表/DB at 84% of one designer's findings and 35% of
another's, behind only 処理フロー・I/O.

**Called from `rev-program-review`?** Then the workbook is already dumped and the scope is already
chosen — don't re-dump it, don't unconditionally launch the other checks, and don't post a
per-check status update. This skill runs on its own when it was one of the selected checks, or when
the user asked for this check by name.

## Environment constraints (important)

**Where the shared docs live:** the `_shared/*.md` files ship **inside this plugin**, in the
`_shared/` folder next to this skill's own directory (`<plugin root>/skills/_shared/`) — **not** in
`~/.claude/skills/_shared/`, which does not exist on a normal install. Resolve every `_shared/...`
reference below against that folder; if it doesn't resolve, glob
`**/rev-skills/**/skills/_shared/<filename>` and read the hit. Do not skip it and improvise the
Excel COM dump instead: that doc carries the guard against attaching to — and then `Quit()`-ing —
the user's own live Excel session, the strikethrough-exclusion cascade, and the reference-file cache.

Read `_shared/xlsx-excel-com-dump.md` first — it has the PowerShell + Excel COM script this skill
(and its sibling review skills) use to read `.xlsx` files, since this machine has no working
Python/Node to use a library like openpyxl. Its "Folders to always exclude from project-wide
searches" list applies here too.

## Sheet anatomy (measured, not assumed)

Verified against all 13 `更新条件表(*)` sheets of
`01_Doc\08_機能定義書\11_工程管理\PHASE3\PXJCO192_処置指示登録.xlsx`. Confirm the shape holds on the
workbook in front of you before trusting the column numbers — the **labels** are the contract, the
column numbers are just the common case.

```
[8,2]=更新ﾃｰﾌﾞﾙ      [8,16]=TXJCM137WF:処置指示ﾏｽﾀWF      ← <TableID>:<table name>
[9,2]=更新概要        [9,16]=①画面(GXJC192A)、②ﾛｸﾞｲﾝ情報   ← defines the ①②… source symbols
[12,2]=更新条件       [12,16]=保存ﾎﾞﾀﾝ押下時 / 分岐条件      ← the trigger, often multi-line
[14,2]=No. [14,4]=項目名 [14,16]=INSERT [14,34]=UPDATE      ← the header row
[15,2]=1   [15,4]=会社ｺｰﾄﾞ [15,16]=② [15,18]=会社ｺｰﾄﾞ  [15,34]=② [15,36]=[KEY]会社ｺｰﾄﾞ
```

Rules that make the parse reliable:

- **The header row is the row with `No.` in col 2 AND `項目名` in col 4.** Do not anchor on row 14.
  Row 13 frequently carries a free-text flow note that happens to contain an operation word
  (`[13,16]=DELETE⇒INSERT`) and will be mistaken for the header by a looser rule.
- **One sheet can hold several blocks.** Half the sheets checked had two or three (e.g.
  `更新条件表(TXJCM137)` has header rows at 14, 68 and 122; `TXJAM026WF` at 14 and 164). Scan the
  whole sheet for header rows and treat each block independently — each has its own 更新ﾃｰﾌﾞﾙ,
  更新条件 and operation set. Stopping at the first block silently skips most of the specification.
- **Operation columns are read off the header row, not hardcoded.** Observed label columns: 16, 28,
  34, 40. Accept a cell as an operation only when its text matches `^(INSERT|UPDATE|DELETE|MERGE)\d*$`
  — suffixed forms (`INSERT1`, `INSERT2`) and repeated plain labels (two `INSERT` blocks in
  `TXJAM026WF`) are both real. This filter matters: several sheets put an unrelated
  取得ﾃｰﾌﾞﾙ/検索条件 sub-block on the same rows further right (`[14,58]=ﾊﾟﾗﾒｰﾀﾏｽﾀ.ｷｰ1`,
  `[14,58]=共通ｺｰﾄﾞﾏｽﾀ.ｷｰ1`), and a "any non-empty cell to the right is an operation" rule turns
  those into phantom operations.
- **Each operation is two sub-columns: the label column `c` holds the source symbol (`①`, `②`, `-`),
  and `c+2` holds the value.** Confirmed for every observed label column (16→18, 28→30, 34→36,
  40→42).
- **"Not set" means the *value* column is `-` or empty — never the source column.** A column set
  from something other than the screen/login sources legitimately reads `[17,16]=-` with
  `[17,18]=ｼｽﾃﾑ日時`. Judging by the source column alone reports every system-supplied column as a
  missing value.
- **`[KEY]` prefixes the value** and marks the column as part of the WHERE clause for
  UPDATE/DELETE (`[15,36]=[KEY]会社ｺｰﾄﾞ`). A literal in a key position is normal
  (`[28,18]=[KEY]"000"`).
- **Struck-through blocks are already gone from the dump.** The shared dump resolves strikethrough
  live, so a withdrawn 更新条件表 block simply won't be there (in `PXJCO192`,
  `更新条件表(TXJCM137)` is a visible sheet whose entire content is struck out and marked
  `2026/6/10 福浦 削除`). Do not resurrect it from `_DELETED_DIGEST.txt` to review it. A table whose
  update spec was withdrawn but which is still declared in Ⅲ．入出力定義 is a real finding — but it
  belongs to `design-doc-io-table-check`, not here.
- **Hidden sheets are out of scope**, per the shared dump doc's default.

## Procedure

### 1. Get the workbook dump

If `rev-program-review` already dumped it, read those scratchpad files. Otherwise dump the workbook
per `_shared/xlsx-excel-com-dump.md`. Only `更新条件表(*)` sheets matter for this check, plus the
機能定義書 and 画面設計書 sheets for the trigger cross-check in step 5 C6.

### 2. Resolve each 更新ﾃｰﾌﾞﾙ to its layout file

The `更新ﾃｰﾌﾞﾙ` cell gives `<TableID>:<name>`. Resolve the ID to exactly one `ﾃｰﾌﾞﾙﾚｲｱｳﾄ` file using
**`xlsx-db-column-check`'s step 3 "Quick reference"** — read that section and follow it rather than
re-deriving it. The rules that bite hardest here: the search root is the top-level
`<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)` (not nested under `01_Doc`), the ID must match
**exactly** (a `WF` suffix is a different table, and `更新条件表` sheets are full of `*WF` tables), the
`ﾃｰﾌﾞﾙﾚｲｱｳﾄ` sheet's `A6` cell must confirm the ID, and precedence is PH3 > PH2 > top-level.

Dump all of them in one pass with the **Batch variant** script in the cross-session cache section of
`_shared/xlsx-excel-com-dump.md`, `OnlySheetPatterns = @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ")`.

The layout sheet's header is at row 7 and the columns this check needs are:

| col | header | use |
|---|---|---|
| 1 | `No.` | column order |
| 3 | `項目名` | the name the 更新条件表 uses |
| 12 | `項目ID` | for reporting |
| 28-35 | `I01`…`I08` | index membership; **`I01` is the primary key** — the digit in the cell is that column's position within the key |
| 36 | `default` | a DB-side default, relevant to C2 |
| 38 | `notnull` | `Y` = NOT NULL |

### 3. Calibrate the house convention from the workbook itself

Before flagging anything about the 共通項目 block, tabulate what **this workbook's own** sheets do for
each common column, per operation. The conventions below were measured on `PXJCO192`, and they held
across its sheets — but they are a project convention, not a written rule, so the workbook's own
dominant pattern wins over this table when the two disagree. Flag a deviation from the workbook's
dominant pattern; do not flag the pattern itself.

Observed standard for the leading common-column block
(`会社ｺｰﾄﾞ/登録者/登録日時/更新者/更新日時/更新ﾎｽﾄ名/更新ﾌﾟﾛｸﾞﾗﾑID/排他ﾌﾗｸﾞ/部門GRP/事業部ｺｰﾄﾞ/移行ｵﾌﾞｼﾞｪｸﾄID`):

| 項目名 | INSERT | UPDATE | DELETE |
|---|---|---|---|
| 会社ｺｰﾄﾞ | 設定 | `[KEY]` | `[KEY]` |
| 登録者 | 作業者ｺｰﾄﾞ | **`-`** | `-` |
| 登録日時 | ｼｽﾃﾑ日時 | **`-`** | `-` |
| 更新者 | 作業者ｺｰﾄﾞ | 作業者ｺｰﾄﾞ | `-` |
| 更新日時 | ｼｽﾃﾑ日時 | ｼｽﾃﾑ日時 | `-` |
| 更新ﾎｽﾄ名 | `-` | `-` | `-` |
| 更新ﾌﾟﾛｸﾞﾗﾑID | 画面ID | 画面ID | `-` |
| 排他ﾌﾗｸﾞ | `1` | `+1` | `-` |
| 部門GRP | 設定 | `[KEY]` | `[KEY]` |
| 事業部ｺｰﾄﾞ | 設定 | `-` | `-` |
| 移行ｵﾌﾞｼﾞｪｸﾄID | `-` | `-` | `-` |

WF tables carry four more after that block — `ﾜｰｸﾌﾛｰID`, `ｸﾞﾙｰﾌﾟID`, `ﾜｰｸﾌﾛｰ申請区分`, `元排他ﾌﾗｸﾞ` —
and the WF sheets in the same workbook are the calibration set for those.

### 4. Run the checks

**C1 — 項目リストが実テーブルと一致しているか.** The 更新条件表's `項目名` rows should be the layout's
`項目名` list, in the layout's `No.` order. Report: a layout column with **no row at all** in the
更新条件表 (the sheet was written against an older table definition — the coder has no instruction
for that column); a row whose 項目名 exists in **no** layout column (usually a renamed column, cite
the nearest layout name); and a row order that diverges from the layout (low confidence on its own —
report only when it coincides with added/removed columns, since it is then evidence the sheet was
patched by hand rather than regenerated).

**C2 — INSERT で notnull 列が未設定.** For every operation whose label starts with `INSERT` (or
`MERGE`): every layout column with `notnull = Y` must have a non-`-`, non-empty value. An unset
NOT NULL column is a guaranteed runtime failure, so this is the highest-severity finding this check
produces. Two exemptions, both verifiable from the layout: the column has a `default` value (col 36),
or it is the target of a DB-side trigger/sequence noted in 備考 (col 40). State the exemption rather
than staying silent — "notnull だが default 設定あり" is useful to the reviewer.

**C3 — 主キーの扱い.** From `I01` (col 28), build the table's primary key in position order.
- INSERT: every PK column must be set. A missing one means the row can't be identified afterwards.
- UPDATE / DELETE: every PK column must carry `[KEY]`. **A partial key is the important finding
  here** — `[KEY]` on 3 of a 5-column PK means the statement updates or deletes a *range* of rows,
  which is nearly always unintended; if it is intended, the 更新条件 text should say so, so check the
  trigger text before flagging and quote it either way.
- A `[KEY]` on a column that is **not** in `I01` is also worth a line: either the doc means a
  non-unique filter (fine, but then see the range warning above) or the PK in the layout is wrong.

**C4 — 共通項目の慣習違反.** Compare each block's common-column rows against the calibration from
step 3. The finding shape that actually shows up: **`登録者`/`登録日時` being set on UPDATE.** This is
not hypothetical — `更新条件表(TXJCM137WF)` in `PXJCO192` sets both on its UPDATE
(`[16,36]=作業者ｺｰﾄﾞ`, `[17,36]=ｼｽﾃﾑ日時(...)`) while every other UPDATE block in the same workbook,
including `TXJCM501`'s and `TXJCM137`'s, correctly leaves them `-`. Overwriting 登録日時 on every
update destroys the record's creation time. Also check: 排他ﾌﾗｸﾞ set to `1` instead of `+1` on UPDATE
(the optimistic-lock counter never advances), and 更新者/更新日時 left `-` on an UPDATE.

**C5 — 排他制御が更新と噛み合っているか.** If the table has an `排他ﾌﾗｸﾞ` column and the block has an
UPDATE, the 更新条件 (or the 機能定義書's processing overview) should describe the optimistic-lock
comparison, not just the `+1`. `design-doc-internal-consistency`'s check 7 asks whether exclusive
control is mentioned *at all* for a program that updates; this check is the column-level follow-up —
report only what that check wouldn't already have said, and say plainly that it's the 更新条件表 side
of the same concern so the reviewer doesn't see it as two separate defects.

**C6 — 更新条件（トリガ）が書かれていて、画面イベントと対応しているか.** An empty `更新条件` cell is a
finding on its own. When it names a trigger (`保存ﾎﾞﾀﾝ押下時`, `一時保存ﾎﾞﾀﾝ押下時`), that event should
exist in the 画面設計書's event list for the screen the 更新概要 names. A trigger naming a button the
screen doesn't have is a stale copy-paste from another program's sheet.

**C7 — 参照元記号が定義されているか.** Every `①②③…` used in a value or source cell must be defined in
that block's own `更新概要` cell. A dangling `⑥` (used in `TXJCM501`'s UPDATE) means the reader can't
tell where the value comes from. Symbols are **block-scoped** — check against the block's own
更新概要, not the sheet's first one.

### 5. Verify before reporting

For each candidate finding, re-read the source cells in the dump. The three things that have
produced wrong findings on this shape of sheet:

- reading the **source** column instead of the value column (`c` instead of `c+2`) and reporting
  every system-supplied column as unset;
- anchoring on row 14 and missing the second and third blocks — or worse, comparing block 1's
  operations against block 2's rows;
- treating a right-hand 取得ﾃｰﾌﾞﾙ/検索条件 sub-block's header cells as operation columns.

## Reporting

Write in the user's language, for the designer who owns the doc.

Group by 更新条件表 sheet, and within a sheet by block when there is more than one. For each finding:
name the table and the column, cite the sheet and cell (`更新条件表(TXJCM137WF) [17,36]`) with a
clickable `[filename](relative/path)` link to the workbook, and state the consequence in one clause
(NOT NULL 違反で登録できない / 登録日時が更新で上書きされる / 主キーの一部しか指定されておらず複数行が
更新される). Order by severity: C2 and C3 partial-key first, then C4, then the rest.

Omit: sheets that checked out clean, a count of how many sheets/columns were compared, and any
narration of which files you dumped or resolved. The one process fact worth a line is a
更新ﾃｰﾌﾞﾙ whose layout file couldn't be resolved at all — that's a coverage gap, and say which table.
