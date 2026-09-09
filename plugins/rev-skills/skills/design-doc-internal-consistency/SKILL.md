---
name: design-doc-internal-consistency
description: Review one program's design-doc set (機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表) for internal cross-reference consistency — e.g. every screen event is reflected in the processing overview, response definitions are registered, every screen-item ID / message ID is registered in its master list, screen layout/item-definition/control-spec agree (including item ORDER matching between Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 and Ⅴ．画面項目定義, not just which items are present), exclusive-control is present when a program updates a table. Grounded in this project's own official review checklist. Does NOT cover the Ⅲ．入出力定義 (I/O table) completeness check — that's the more involved `design-doc-io-table-check` skill, split out separately. Use when the user asks to レビュー/REV a program's 機能定義書 or 画面設計書, or asks whether a design doc is "internally consistent" / "漏れがないか", or whether screen-item ordering matches across sections. When the user asks to REV a single design-doc workbook without naming which checks they want, the entry point is `rev-program-review`: it first asks the user, checkbox-style, which of the 6 single-program checks to run, then runs only those as one combined pass. Do not launch all six yourself. Run this skill standalone only when it was one of the selected checks, or when the user asked for this check by name.
---

# design-doc-internal-consistency

Reviews a single program's design-doc workbook (機能定義書 + 画面設計書 + ﾁｪｯｸ処理設計書 +
更新条件表, usually all sheets inside one `<ProgramID>_<name>.xlsx`) for cross-reference
consistency, grounded in this project's own checklist:
`01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx` (sheet
"ﾁｪｯｸﾘｽﾄ"). Read that sheet at the start of a review — it is short (~57 rows) and is the
authoritative list of what "correct" means here; the checks below are the checklist's
programmatically-verifiable items 3-4 through 5-4 (item 3-2/3-3, the Ⅲ．入出力定義 completeness
check, now lives in the separate `design-doc-io-table-check` skill — see below), restated as
concrete diffing steps. The checklist also has purely-manual/subjective items (e.g. "列幅は2、
行幅は14.3か" 1-4) — skip those, they aren't worth automating.

This skill checks **cross-reference consistency** only (does X mentioned here also appear over
there) — see the frontmatter `description` above for exactly which adjacent checks (I/O table
completeness, DB-column existence, ID-numbering, formatting, typos) live in the 5 sibling skills
instead.

**Scope selection comes first.** When the user asks to REV a single program's design-doc workbook
without naming specific checks, `rev-program-review` is the entry point: it presents the 6
single-program checks as a checkbox list (`AskUserQuestion`, multiSelect), then runs only the
selected ones as one combined pass, dumping the workbook once up front and sharing the text with
every check (see `_shared/xlsx-excel-com-dump.md`'s "dump once, share the text" section). Do **not**
unconditionally launch all 6 yourself, and do not post a per-check status update — the combined
report is posted once, after every selected check has finished. This skill runs on its own when it
was one of the selected checks, or when the user asked for this check by name.

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
struck-through / grayed-out rows from review" section. **The dump script already dropped
struck-through and grayed-out content, so do NOT run a formatting scan of your own** — every row you
can see is live, which is what makes the "used but not declared" / "missing from dictionary" diffs
below trustworthy without any extra work.

**Of the six single-program checks, this is the one that genuinely needs `_DELETED_DIGEST.txt`.**
Its highest-value findings are asymmetries — one half of a pair edited while the other was left
behind — and the live dump only ever shows you the survivor. Read the digest when a check turns on
what was *removed*:

- a ※-note, a branch table row, or a step number whose counterpart was deleted, leaving a dangling
  definition or a hole in a sequence;
- prose that still advertises an output/behaviour whose implementing rows were deleted (confirmed on
  `SXJCB147`: Ⅰ．機能概要 still said 識別ｶｰﾄﾞ was printed after every 識別ｶｰﾄﾞ output row was struck);
- a "削除" revision note sitting beside content that is *not* struck — genuinely ambiguous, and worth
  reporting as a judgment call rather than resolving yourself.

Conversely, do not report something as "a leftover that should have been deleted" without checking
the digest first: it may already be marked dead, in which case there is nothing to fix. That exact
false finding has been made before (`XJC_ｼｽﾃﾑ共通設計書.xlsx`, `ﾛｯﾄ停止ﾁｪｯｸ`, rows 829-832).

**Don't dump or read `詳細設計書` sheets.** None of the 7 checks below reference detail-design
content — they all draw from 機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表 only. Across every
program reviewed with this skill so far, 詳細設計書 has turned out to be a near-empty
header-only sheet anyway, so skipping it is a pure token saving with nothing lost. If a future
review genuinely needs detail-design content for some other purpose, that's a different task outside
this skill's scope, not a reason to dump it here by default.

## Checks to run (checklist item → what to diff)

For each check, dump the relevant sheets (per the shared doc) and read them, then compare. Cite the
exact sheet/cell for every finding so it's actionable.

1. **3-4 — Event ↔ processing-overview completeness.**
   Collect every event listed in 画面設計書 "Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細" (button clicks, フォーカスロスト,
   起動時, etc. — one block per screen area). Collect every step described in 機能定義書
   "Ⅳ．機能処理概要". Every event block in Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 should have a corresponding step in
   機能処理概要 (they don't need identical wording, but the described behavior should be traceable).
   Flag any event with no corresponding processing-overview step.

2. **3-5 — Response definitions registered.**
   If 機能定義書 "Ⅶ．ﾚｽﾎﾟﾝｽ定義" defines anything beyond "※ｼｽﾃﾑ共通設計書.ﾚｽﾎﾟﾝｽ 参照", confirm it's
   reflected in `01_Doc/06_システム設計書（一覧、管理台帳）/06-09_ﾚｽﾎﾟﾝｽ一覧_共通.xlsx`.

3. **4-12/7-3/8-3 — Screen-item IDs registered in the dictionary.**
   Collect every 画面項目ID from 画面設計書 "Ⅴ．画面項目定義" (rightmost ID column, e.g. `XJC0036`,
   `SJC0662`) — note the ID is usually split across two adjacent cells in the dictionary/master
   files (a 3-letter prefix column, then a 4-digit number column) even though it appears as one
   token in the screen design sheet; concatenate them when extracting from the master.
   **Each ID routes to a dictionary file by its own prefix, not by which WG is under review** —
   `XJZ`/`SJZ` → `_共通`, `XJA`/`SJA` → `_基準情報`, `XJB`/`SJB` → `_受注出荷`, `XJC`/`SJC` →
   `_工程管理`, `XJD`/`SJD` → `_品質管理`, all under `01_Doc/04_共通設計/`; a `MENU…` ID routes on the
   prefix that follows `MENU`. So a 工程管理 program that uses a shared item is checked against the
   `_共通` file for that ID, and checking it against `_工程管理` would report a registered item as
   unregistered. `_shared/reference-index.md` carries the full table and the reason an ID is never
   satisfied by a same-named sheet embedded in another WG's copy. Expect to be handed one index per
   file the program's IDs reach — usually one or two. Flag any ID used in the screen design but
   missing from whichever file should hold it (or vice versa if the dictionary shows an entry whose
   画面項目名 disagrees with the screen design's 画面項目名 for the same ID — a rename that wasn't
   propagated).

   **Don't dump `82.画面項目辞書_<WG名>.xlsx` — read an index of it. See
   `_shared/reference-index.md`.** This check needs only `id → 画面項目名`, and an index of that costs
   one to two orders of magnitude fewer tokens than the dump this step used to take, less again once
   subset to the IDs one program actually cites; the figures and the source state they were measured
   at are in that doc, and are deliberately not repeated here. It also carries the two structural
   traps this step kept hitting: the ID a design doc cites is the `ＩＤ` and `連番` sub-columns
   **joined**
   (`XJC` + `8046` = `XJC8046`), so reading either alone yields IDs that appear in no design doc at
   all; and the dictionary sheets have to be located by their `画面項目ID`/`画面項目名` header rather
   than by sheet name, because every WG's copy names them differently and some copies carry
   non-dictionary sheets under other WGs' names.

   The index is built by the orchestrating session before agents launch, not by you. If you were
   pointed at one that is missing or stale, say so rather than building it or opening Excel.

   Work from the index for both directions of this check: an ID cited by the doc but absent from the
   index is unregistered, and a 画面項目名 that differs from the index's is the rename-not-propagated
   finding. Entries retired by strikethrough are excluded from the index, so an ID present in it is
   genuinely registered — but an entry **renamed in place** (old name struck, new name live in the
   same cell) is kept, carrying its live name, so a name mismatch against one of those is a real
   finding, not an artefact.

   You are normally handed **two** paths: a *subset* of the index scoped to this program, to read
   whole, and the *full* index, to grep. The subset is the full index filtered down to the IDs found
   in this program's own dump — so an ID this doc cites that is absent from the subset already **is**
   the unregistered-ID finding. Do not dismiss it as "the subset just doesn't cover it". Confirm it
   with one grep of the full index before reporting: absent from both is the finding; present in the
   full index means the subsetting filter missed that ID's form, which is worth saying but is not a
   design defect.

4. **4-8/5-1 — Message IDs registered.**
   Collect every message ID referenced in 画面設計書 "Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細" (ﾒｯｾｰｼﾞ column) and in
   ﾁｪｯｸ処理設計書 (error codes) — format is `<3-char JOBｺｰﾄﾞ>-<0/1><5-digit>`, e.g. `XJC-000131`.
   The middle letter of the JOBｺｰﾄﾞ identifies which WG owns that prefix — `JA`=基準情報,
   `JZ`=共通, `JC`=工程管理, `JB`=受注出荷, etc. (extend this mapping as you encounter other WG
   prefixes). Route the lookup by that owning WG, not by which program you're reviewing:
   - **Program's own-WG prefix IDs** (e.g. an `XJC`/`SJC` id inside a 工程管理 program): check
     against that WG's own `01_Doc/04_共通設計/04.ﾒｯｾｰｼﾞ管理_<WG名>.xlsx` (e.g.
     `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx`). **Dump it via the cross-session cache with
     `OnlySheetPatterns = @("<X-prefix>(*", "<S-prefix>(*")`** (e.g. `@("XJC(*", "SJC(*")` for
     工程管理, open-ended with no closing `)` — same reason as the screen-item dictionary case
     above) instead of a full-file dump: confirmed for real that
     `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx` has 11 sheets but only the program's own two prefix sheets are
     ever read from it — a foreign-WG prefix routes to that other WG's own file (see below), never
     to a same-named sheet embedded in this one, so the other 9 sheets in this file are dead weight
     for every 工程管理-program review. See `_shared/xlsx-excel-com-dump.md`'s cross-session cache
     section for the full mechanism.
   - **`XJZ`/`SJZ` (共通) prefix IDs, from *any* program in *any* WG**: check ONLY against
     `01_Doc/04_共通設計/04.ﾒｯｾｰｼﾞ管理_共通.xlsx` (sheets `XJZ(共通)`/`SJZ(共通)`). **Do not use a
     WG-specific file's own embedded `XJZ(共通)`/`SJZ(共通)` sheet as the source of truth for
     this** — those per-WG copies have been observed to be stale duplicates that drift from the
     real common file (e.g. `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx`'s `XJZ(共通)` sheet showing an outdated
     wording for `XJZ-000020` while the real `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` has the current, correct
     text that actually matches every design doc). Checking the wrong copy produces a **false
     "text mismatch" finding** — this happened repeatedly across several program reviews before
     being caught. If a WG-local `XJZ(共通)`/`SJZ(共通)` sheet disagrees with the real common
     file, that staleness in the WG-local copy is not itself worth reporting (it isn't the design
     doc under review) — just silently use the common file as ground truth.
   - **A prefix belonging to a WG other than the one you're reviewing and other than 共通** (e.g. an
     `SJA-xxxxx` id inside a 工程管理 program): don't assume it's simply missing — note it
     separately as "cross-WG reference, couldn't verify without checking that other WG's own
     message file" rather than a confirmed gap.
   - **A short, unhyphenated `<3-char JOBｺｰﾄﾞ><4-digit>` reference (e.g. `XJZ7006`, `XJC7006`,
     `SJC7013`) — most often in 画面設計書 Ⅳ's own ﾒｯｾｰｼﾞ column labeled "処理前:"/"完了後:" pairs for
     a confirm/complete dialog — is a DIFFERENT ID namespace from the hyphenated
     `<JOBｺｰﾄﾞ>-<6-digit>` error-message IDs above, and lives in a different file.** Confirmed for
     real on `PXJCO124_ﾛｯﾄ振向け.xlsx`, `画面設計書(GXJC124A)` row 807 (実行ﾎﾞﾀﾝ: 処理前 `XJZ7006`/
     完了後 `XJC7006`) and row 810 (再印刷ﾎﾞﾀﾝ: `SJC7013`) — none of these exist anywhere in
     `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` or `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx` (their `000001～`/`100001～`/`200001～`
     numbering blocks don't even reach a `7xxxxx` range), which looks like a registration gap at
     first — but all three IDs turned out to be genuinely registered as `MESSAGE`-category rows
     inside the **screen-item dictionary** files instead: `XJZ7006` in
     `82.画面項目辞書_共通.xlsx`'s `XJZ(共通_JAGUR)` sheet, `XJC7006`/`SJC7013` in
     `82.画面項目辞書_工程管理.xlsx`'s `XJC(工程)`/`SJC(工程)再開発追加分` sheets — each with a
     `カテゴリ`/`localizationid` of `'MESSAGE'` even though the file's overall name says "画面項目".
     When a message-shaped ID doesn't hyphenate and doesn't resolve in the `04.ﾒｯｾｰｼﾞ管理_*` files,
     check the relevant `82.画面項目辞書_*` file's dictionary sheets (the `ＩＤ`+`連番` split, same as
     a normal screen-item ID lookup) before concluding it's unregistered.

5. **4-9/4-13 — Screen layout ↔ item-definition ↔ control-spec three-way match.**
   Collect item names from "Ⅰ．画面ﾚｲｱｳﾄ" (the visual mock), "Ⅴ．画面項目定義" (the item table), and
   "Ⅵ．画面項目制御・出力仕様" (the control matrix). All three should list the same set of
   non-hidden items. Flag any item present in one but missing from another (hidden items are
   expected to be absent from Ⅵ per the checklist note — don't flag those).

   **Also check the ○/× marker character used throughout "Ⅵ．画面項目制御・出力仕様" against the
   sheet's own legend line** (e.g. "ｲﾍﾞﾝﾄによる項目制御(使用可/不可、編集値) ※"○":使用可､"×":使用不可"
   — one such legend line typically precedes each item-control block). The legend defines the
   "usable" marker as `○` (U+25CB WHITE CIRCLE), but individual matrix cells sometimes instead use
   `〇` (U+3007 IDEOGRAPHIC NUMBER ZERO / kanji "zero") — a different Unicode codepoint that renders
   almost identically and is easy to type by mistake (e.g. via IME kanji conversion of "まる"), so a
   plain visual read of the sheet won't catch it. Confirmed for real on
   `PXJCO152_処置指示発行.xlsx`, sheet `画面設計書(GXJC152A)`: the legend at row 786/803 defines `○`,
   but cells `[649,33]`, `[792,16]`, `[792,28]`, `[793,28]`, `[806,12]`, `[806,28]`, `[806,32]`,
   `[808,12]`/`[808,28]`/`[808,32]` (and likely more throughout the sheet) use `〇` instead. Compare
   each marker cell's character by Unicode codepoint (not just by eye or by a plain text-equality
   check that might normalize them) against the legend's own `○`, and flag every cell using the wrong
   codepoint as a typo to correct to `○`.

   **Also check the letter-casing of any date/datetime format string in "Ⅴ．画面項目定義"'s 表示形式
   column.** This project's screen-side (client) display-format convention is `yyyy/MM/dd` — lowercase
   `y` for year, uppercase `MM` for month, lowercase `dd` for day (the standard Java
   `SimpleDateFormat`-style casing, where case distinguishes month `MM` from minute `mm`). A
   same-looking but wrong-cased variant, most often fully-uppercase `YYYY/MM/DD`, is a real defect,
   not a stylistic choice — confirmed for real on `PSJCO304_着手ﾒｯｾｰｼﾞﾒﾝﾃﾅﾝｽ.xlsx`, sheet
   `画面設計書(GSJC304A)`, rows `353`/`355`/`361`/`363` (TextBox-type 表示開始日/表示終了日 items) and
   rows `422`/`423` (their Label-type display counterparts in the results list) — all six show
   `YYYY/MM/DD` in column 25 (表示形式) where `yyyy/MM/dd` was intended. **Scope this check to
   画面項目定義's 表示形式 column specifically** — a date-format string appearing elsewhere, e.g. inside
   a 更新条件表's 取得内容 column describing how a value is formatted at the DB/SQL layer (e.g.
   `ｼｽﾃﾑ日時(YYYY/MM/DD HH24:MI:SS形式)`), follows Oracle's own `TO_CHAR` format-model convention
   instead (where `HH24`/`MI`/`SS` are themselves Oracle-specific tokens, case-insensitive in that
   context) — don't flag those, they're a different convention entirely and casing carries no meaning
   there. Flag every 表示形式 cell using `YYYY` and/or `DD` (or any other wrongly-cased token relative
   to the `yyyy/MM/dd` convention) as a casing defect to correct.

6. **Screen-item ordering consistency between Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 and Ⅴ．画面項目定義.**
   Collect the order in which items appear in 画面設計書 "Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細" (each screen-area
   block lists its items' event behavior in a fixed sequence) and, separately, the order they
   appear in "Ⅴ．画面項目定義" (the item table, top-to-bottom, normally following 画面ﾚｲｱｳﾄ/tab
   order). Both sections describe the same set of items and should present them in the same
   relative order within a matching scope (same screen area / same item group) — a reordering
   between the two is a strong signal that one section was edited (an item added, removed, or
   moved) without the other being kept in sync, which often also means the underlying change wasn't
   fully propagated to both places. Flag any pair of items whose relative order differs between the
   two sections (item X appears before item Y in one section but after Y in the other). Exclude
   hidden items from this comparison the same way check 5 does (they're not always listed in every
   section). When you can't tell which of the two orderings is authoritative (neither section
   carries an explicit "this is the source of truth" marker), report the finding as a plain "order
   mismatch" between the two locations without asserting which side is wrong, and let the designer
   decide which one to resequence.

7. **5-4 — Exclusive-control checks present when updating.**
   If any 更新条件表 sheet exists for this program (i.e. it writes to a table), confirm
   ﾁｪｯｸ処理設計書 mentions a 排他ﾁｪｯｸ, and confirm "Ⅴ．画面項目定義" has a hidden 排他ﾌﾗｸﾞ item. Flag
   if a program updates a table but has neither.

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on.

For each real finding: state the defect in one or two sentences, cite the exact sheet name and cell
coordinates (and the master file it was checked against, e.g. 82.画面項目辞書, 04.メッセージ管理,
06-09_レスポンス一覧) so it's actionable, and give a concrete suggested fix when one is
obvious (e.g. "change X to Y"). Group findings under a checklist-item
heading (e.g. "3-4: ...") only when several findings share one; don't force every finding into the
checklist's numbering if it doesn't fit cleanly.

Order by impact, most important first — a broken cross-reference, a missing table, or a functional
gap belongs well before a stray full-width/half-width punctuation difference or a naming nit. When a
finding depends on a judgment call rather than a clear rule (a naming convention that might be
intentional, a cross-WG reference you couldn't fully verify, a pattern that could be deliberate
delegation to a common design doc rather than an omission), say so explicitly and mark it as needing
the designer's confirmation — don't present it with the same confidence as a confirmed defect.

Omit entirely, unless the user explicitly asked for a full pass-by-pass audit:
- "確認済み・問題なし" lines for checks that simply passed
- tallies of how many items/tables/IDs were checked
- narration of your own process (which files you dumped, how you sampled, which scratch path you
  used)

The one exception worth keeping brief: if a whole section couldn't be verified at all (a master file
was missing, a section was empty in this program's doc), a one-line note is fine — that's a real gap
in review coverage the designer should know about, not a process detail.
