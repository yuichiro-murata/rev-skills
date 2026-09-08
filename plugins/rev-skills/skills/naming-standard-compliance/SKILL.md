---
name: naming-standard-compliance
description: Check that program/screen/table/file/report/zoom/message IDs and design-doc headers in this project follow the official ID-numbering rules and doc-writing checklist (05.システム共通設計書「各ID採番」, 05.設計書記述ルール_チェックリスト). Use when the user asks to check ID命名規則/採番ルール compliance, or wants a broad "is this design doc written correctly" pass distinct from cross-reference or DB-column checks. When the user asks to REV a single design-doc workbook without naming which checks they want, the entry point is `rev-program-review`: it first asks the user, checkbox-style, which of the 6 single-program checks to run, then runs only those as one combined pass. Do not launch all six yourself. Run this skill standalone only when it was one of the selected checks, or when the user asked for this check by name.
---

# naming-standard-compliance

Checks that IDs used in design docs (and the docs' own header/structure) follow this project's
documented rules — see the frontmatter `description` above for which sibling skill covers
cross-references, I/O table completeness, DB-column existence, or typo proofreading instead.

**Scope selection comes first.** When the user asks to REV a single program's design-doc workbook
without naming specific checks, `rev-program-review` is the entry point: it presents the 6
single-program checks as a checkbox list (`AskUserQuestion`, multiSelect), then runs only the
selected ones as one combined pass, dumping the workbook once up front and sharing the text with
every check (see `_shared/xlsx-excel-com-dump.md`'s "dump once, share the text" section). Do **not**
unconditionally launch all 6 yourself, and do not post a per-check status update — the combined
report is posted once, after every selected check has finished. This skill runs on its own when it
was one of the selected checks, or when the user asked for this check by name.

Grounded in:

- `01_Doc/04_共通設計/05.ｼｽﾃﾑ共通設計書.xlsx`, sheet **"各ID採番"** — the numbering rule for each ID
  type (read it fresh each run; do not rely on a paraphrase from memory, since it's the single
  source of truth and this skill's job is to catch drift from it).
- `01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx`, sheet
  **"ﾁｪｯｸﾘｽﾄ"** — items 1-x/2-x (general/表紙 formatting rules) plus the header-info check (1-6:
  機能ID・システムID・作成日 etc. must be correct and consistent across every sheet of one
  workbook).

## Environment

**Where the shared docs live:** the `_shared/*.md` files ship **inside this plugin**, in the
`_shared/` folder next to this skill's own directory (`<plugin root>/skills/_shared/`) — **not** in
`~/.claude/skills/_shared/`, which does not exist on a normal install. Resolve every `_shared/...`
reference below against that folder; if it doesn't resolve, glob
`**/rev-skills/**/skills/_shared/<filename>` and read the hit. Do not skip it and improvise the
Excel COM dump instead: that doc carries the guard against attaching to — and then `Quit()`-ing —
the user's own live Excel session, the strikethrough-exclusion scan, and the reference-file cache.

Read `_shared/xlsx-excel-com-dump.md` first for how to dump `.xlsx` sheets to text
via PowerShell + Excel COM (no Python/Node available here) — including its "Excluding
struck-through / grayed-out rows from review" section. An ID string that's struck through or grayed
out in the source sheet is marked for deletion, not a live ID to validate — run that section's
targeted formatting scan before extracting IDs, and drop any flagged cell rather than checking it
against the numbering rule.

## ID rules (from "各ID採番" — re-verify against the live sheet, this is a summary)

| ID種別 | 構成 | 例 |
|---|---|---|
| プログラムID | `[P=プログラム\|S=サブプログラム]` + `<3桁JOBコード>` + `[O=オンライン\|B=バッチ]` + `<3桁連番(0埋め)>` | `PSJCO315`, `PSJCB315` |
| 画面ID | `G` + `<3桁JOBコード>`(プログラムIDと同一) + `<3桁連番>`(プログラムIDと同一) + `<プログラム内連番:A,B,C...>` | `GSJC315A` |
| テーブルID | `[T=テーブル\|V=ビュー]` + `<3桁JOBコード>` + `[M=マスタ\|D=データ\|V=トラン\|A=累積\|P=パラメータ]` + `<3桁連番(0埋め)>` | `TXJAM025`, `TSJCV501`, `VSJAM060` |
| ファイルID | `F` + `<3桁JOBコード>` + `<3桁連番(0埋め)>` | `FSJA004` |
| 帳票ID | `R` + `<3桁JOBコード>` + `<3桁連番(0埋め)>` | — |
| ズームID | `Z` + `<基準情報の固定JOBコード>` + `<3桁連番(0埋め)>` | `ZXJA047` / `ZSJA047` — the rule sheet's literal text only shows `XJA`, but in practice **both `XJA` and `SJA` are valid**, following the same 共通(X)/個社固有(S) pairing used by every other ID type's JOBコード in this project (e.g. `XJC`/`SJC` for プログラムID). Confirmed by the user directly — this is not a rule violation, don't flag `SJA` zoom IDs at all. Only flag a zoom ID whose basic-info segment is neither `XJA` nor `SJA`. |
| メッセージID | `<3桁主管理JOBコード>` + `-` + `[0=バインド変数無し\|1=バインド変数有り]` + `<5桁連番(0埋め)>` | — |
| 区分コード | `<3桁主管理JOBコード>` + `<5桁連番(0埋め)>` | — |
| 共通コード | `<3桁JOBコード>` + `<キー1:2桁>` + `<キー2:90桁>` + `<キー3:150桁>` + ボディ(30項目×各1000桁) | — |

Note: **画面項目ID** (e.g. `XJC0036`, `SJC0662` seen inside 画面設計書「Ⅴ．画面項目定義」) is a
*different* ID space not defined in "各ID採番" at all — its correctness is "is it registered in
`82.画面項目辞書_共通.xlsx`", which is `design-doc-internal-consistency`'s job, not a fixed regex
to validate here.

## Procedure

1. Given a target file, WG folder, or the whole project, enumerate design-doc workbooks in scope
   (`表紙`/`機能定義書`/`画面設計書`/`ﾃｰﾌﾞﾙﾚｲｱｳﾄ` sheets carry the IDs — dump those sheets per the
   shared COM script).
2. Extract every ID of each type from its home location:
   - プログラムID/画面ID from the 表紙 and 機能定義書/画面設計書 header rows (row 3-4 area, labeled
     `ﾌﾟﾛｸﾞﾗﾑID`/`画面ID`).
   - テーブルID from ﾃｰﾌﾞﾙﾚｲｱｳﾄ row 6.
   - ファイルID/帳票ID/ズームID from wherever the doc set defines them (ﾌｧｲﾙ出力仕様書/帳票設計書/
     ズーム設計書 headers, or references to them inside 画面設計書 event descriptions).
3. Parse each extracted ID against its rule's regex shape above. Flag:
   - IDs that don't parse at all (wrong length, unexpected letter in a fixed-value slot).
   - IDs where the JOBコード segment doesn't match the JOBコード actually used elsewhere in the same
     doc (e.g. プログラムID says `SJC` but 画面ID says `SJA` for the supposedly-paired screen —
     these two IDs are defined by the rule to share the same JOBコード and sequence).
   - Object-type-letter mismatches for テーブルID (e.g. ID starts with `V` but the workbook's own
     ﾃｰﾌﾞﾙ名/構造looks like a physical table, not a view, or vice versa).
4. Header-info check (checklist 1-6): within one workbook, confirm 機能ID/画面ID/システムID/
   計画書No/作成日 are identical across every sheet's header block (表紙, 機能定義書, 画面設計書,
   ﾁｪｯｸ処理設計書, 更新条件表 all repeat this header — they should never disagree). Flag any sheet
   whose header contradicts the others in the same file. **Deliberately skip `詳細設計書` sheets
   here** (and don't dump them at all) — across every program reviewed with this skill's sibling
   `design-doc-internal-consistency` so far, 詳細設計書 has turned out to be a near-empty
   header-only sheet, so checking its header costs a full-sheet dump for essentially no chance of
   catching a real mismatch. This is a deliberate, documented coverage trade for token savings, not
   a silent gap — if the user specifically asks about 詳細設計書 header consistency, check it then.

   **Also check WITHIN each individual sheet, not just across sheets**: every header block in this
   project's template is physically duplicated side-by-side on one row (e.g. row 1's `2C` label
   with the header starting around column 5, and a second `2D`-labeled copy of the same header
   starting around column 57 — the same pairing exists for `3C`/`3D` on 画面設計書/ﾁｪｯｸ処理設計書/
   更新条件表/etc.). These two copies are meant to be identical but can drift independently of the
   cross-sheet check above. This happened for real on `PXJCB102_製造ｵｰﾀﾞｰ完了処理.xlsx`: both
   `機能定義書(PXJCB102)` and `更新条件表(TXJCM003)` showed 作成日=`46197` in the left (`2C`/`3C`)
   copy but 作成日=`-` (blank) in the right (`2D`/`3D`) copy of the very same header row — while a
   third sheet in the same workbook, `ﾁｪｯｸ処理設計書(PXJCB102)`, correctly showed `46197` on both
   sides. A cross-sheet-only check would have missed this, since it only compares one canonical
   value per sheet and never looks at whether a sheet agrees with itself. Extract both copies from
   every sheet's header row and diff them against each other in addition to the cross-sheet diff.

5. **表紙's "Ⅱ．設計書構成" list vs the sheets actually present** (checklist item near 2-1) — run
   this as a **standard, always-on check**, not an optional one: compare every row's Customer/
   Developer marker (`●`/`-`) against whether a matching sheet genuinely exists in the workbook
   *and has real content* (not just an empty template). This has caught a real defect in every
   program reviewed with it so far that had extra doc types beyond the bare minimum: on
   `PXJCB134_流動停止(ﾊﾞｯﾁ).xlsx`, 表紙 marked both `ﾁｪｯｸ処理設計書`(No.4) and `詳細設計`(No.9) as
   `-`/`-` (not applicable) despite both `ﾁｪｯｸ処理設計書(PXJCB134)` (populated with real check
   items) and two detail-design sheets existing — and 表紙's own 改訂履歴 No.4 even said "詳細設計書
   を作成" (created the detail-design doc), directly contradicting the `-` marker. On
   `PXJCB102_製造ｵｰﾀﾞｰ完了処理.xlsx`, the same thing happened for `詳細設計`(No.9) alone. Given this
   hit rate, don't treat this as a "only if asked for a full pass" item — check it every time, the
   same as the header-info check above.

6. 表紙 checks that remain lighter-weight/manual-leaning (do these only if asked for a full pass,
   they require reading prose or a slower cross-file lookup): 表紙's 画面ID/画面名 matches the
   program's WG-specific 機能一覧 workbook under
   `01_Doc/06_システム設計書（一覧、管理台帳）/06-01_機能一覧_<WG名>.xlsx` (e.g.
   `06-01_機能一覧_工程管理.xlsx` for a 工程管理 program, `06-01_機能一覧_品質管理.xlsx` for
   品質管理, etc.) — **not** `06-01_機能一覧_共通.xlsx`. That "共通" workbook only lists the
   cross-cutting `XJZ`-prefixed common functions (confirmed by dumping it: its one data sheet is
   named `共通`, ~52 rows, all `機能分類=共通(XJZ)`) and will not contain a program from any other
   WG at all — a program ID/機能ID lookup there for e.g. a `SJC`-prefixed 工程管理 program simply
   comes back "not found", which is a wrong-file miss, not a real "画面ID missing from 機能一覧"
   finding. Pick the 機能一覧 file whose WG suffix matches the target program's JOBコード domain
   (`XJC`/`SJC` → 工程管理, `XJB`/`SJB` → 受注出荷, etc.), and only fall back to `_共通.xlsx` when
   the program itself is genuinely an `XJZ` common one. In the WG-specific sheet (sheet name matches
   the WG, e.g. `工程管理`), the row for a given `PGMID` (around col 34) also carries the screen's
   画面/帳票/ﾌｧｲﾙID (col 49, header `画面/帳票/ﾌｧｲﾙID`) and a 処理内容 label (col 53) whose first
   line is the screen name — that's the pair to diff against the workbook's own 画面ID/画面名, since
   this sheet has no separate dedicated 画面ID column of its own. Also (still in this lighter
   category): new files have a 改訂履歴 entry marked 新規作成 (and 流用新規 files cite their 流用元
   計画書No + 改訂履歴 No).

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on.

Group by ID type / checklist item, but only for types that actually produced a finding. For every
flagged ID, show the raw string, which rule it was checked against (quote the rule briefly), the
location it was found, and — when the fix is obvious — what it should be instead. Order by impact:
a JOBコード mismatch between paired IDs or a header-info contradiction (checklist 1-6) belongs before
a cosmetic 表紙 formatting nit. Distinguish clear rule violations from "the rule sheet and reality
disagree, worth asking a human" cases and mark the latter as needing the designer's judgment — don't
assert the design doc is wrong when it's equally possible the common-design rule sheet is the stale
one (though note the ズームID `XJA`/`SJA` case above is now a *resolved* example of this, not an
open one — don't re-flag it).

Omit entirely: IDs that parsed cleanly and matched (don't list "プログラムID: OK"), a tally of how
many IDs of each type were checked, and narration of your own process. The one exception worth a
one-line mention: an ID category you couldn't check at all because its home sheet/section was
missing or empty — that's a coverage gap worth flagging, not a process detail.
