---
name: design-doc-typo-check
description: Proofread the free-text Japanese prose inside a program's design-doc workbook — feature/processing narratives, event descriptions, error-message text, footnotes, revision-history notes — for actual language mistakes (誤字/脱字/衍字/助詞の誤り/変換ミス), as distinct from its sibling skills' structural checks: cross-reference completeness (design-doc-internal-consistency), I/O table completeness (design-doc-io-table-check), DB column existence (xlsx-db-column-check), ID-numbering rules (naming-standard-compliance), and font/merge irregularities (design-doc-formatting-consistency). Use when the user asks to check 誤字脱字 or wants a language-proofreading pass on a design doc. When the user asks to REV a single design-doc workbook without naming which checks they want, the entry point is `rev-program-review`: it first asks the user, checkbox-style, which of the 6 single-program checks to run, then runs only those as one combined pass. Do not launch all six yourself. Run this skill standalone only when it was one of the selected checks, or when the user asked for this check by name.
---

# design-doc-typo-check

Proofreads one program's design-doc workbook for genuine Japanese-language errors in its free-text
prose — wrong kanji, a missing character that breaks the sentence, a stray duplicated character, a
wrong particle that flips the meaning, a garbled leftover fragment from a copy-paste edit. This is
fundamentally different from every sibling REV skill: those are mechanical diffing/lookup tasks;
this one requires actually reading each sentence and understanding what it says. Do not try to
shortcut it with keyword search or pattern matching — read the prose.

This skill does **not** check cross-document terminology consistency, ID format, or font/merge
signals — see the frontmatter `description` above for which sibling skill covers each. A formatting
anomaly `design-doc-formatting-consistency` flags is often a good hint of *where* a copy-paste-era
typo might also be hiding, though, so it's worth cross-referencing that skill's findings if run in
the same pass.

**Scope selection comes first.** When the user asks to REV a single program's design-doc workbook
without naming specific checks, `rev-program-review` is the entry point: it presents the 6
single-program checks as a checkbox list (`AskUserQuestion`, multiSelect), then runs only the
selected ones as one combined pass, dumping the workbook once up front and sharing the text with
every check (see `_shared/xlsx-excel-com-dump.md`'s "dump once, share the text" section). Do **not**
unconditionally launch all 6 yourself, and do not post a per-check status update — the combined
report is posted once, after every selected check has finished. This skill runs on its own when it
was one of the selected checks, or when the user asked for this check by name.

## Environment

Read `~/.claude/skills/_shared/xlsx-excel-com-dump.md` first for how to dump `.xlsx` sheets to text
via PowerShell + Excel COM (no Python/Node available here) — including its "Excluding
struck-through / grayed-out rows from review" section. Struck-through/grayed text is marked for
deletion, not live prose — run that section's targeted formatting scan on any block before
proofreading it, and silently skip flagged rows (don't proofread deleted-in-spirit text, and don't
count skipped rows as findings). As with the sibling skills, **don't dump or read `詳細設計書`
sheets** — out of scope for the same token-saving reasons documented in the other REV skills.

## Scope: which cells are "prose" here

Structured tokens — IDs, table/column names, `<alias>.<column>` expressions, numeric codes, half-width
katakana terms that are this project's standard house style (`ﾃﾞｰﾀ`, `ﾛｯﾄ`, `ｵｰﾀﾞｰ`, `ｺｰﾄﾞ`, etc.) —
are NOT prose and are NOT this skill's concern; those belong to the other REV skills or are simply
normal project vocabulary. Only free-running Japanese sentences and phrases are in scope. In a typical
program workbook, that means:

- **表紙**: Ⅲ．改訂履歴 の 改訂内容 column (what each revision actually changed).
- **機能定義書**: Ⅰ．機能概要 (prose bullets), Ⅱ．機能要求事項, Ⅲ．入出力定義 の 用途 column
  (short prose), Ⅳ．機能処理概要 (the main narrative — numbered steps, ※ footnotes, and any prose
  woven into 検索条件/取得内容 descriptions), Ⅵ．前提条件, Ⅷ．その他特記事項.
- **画面設計書**: Ⅲ．画面表示仕様 の ※ footnotes, Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 (event-behavior
  descriptions — often long paragraphs), Ⅴ．画面項目定義 の 備考 column, Ⅵ．画面項目制御・出力仕様
  の 備考/条件 prose.
- **ﾁｪｯｸ処理設計書**: condition descriptions, error-message body text.
- **更新条件表**: the 更新概要 prose at the top of the sheet, and ※-numbered footnotes (these are
  written in full sentences, unlike the 項目名/取得内容 columns which are mostly structured tokens).
- **帳票設計書** (if present): 処理の流れ (numbered narrative steps), 備考 column.

## Procedure

1. Dump the target workbook's sheets per the shared COM technique (skip 詳細設計書).
2. For each section listed above, read every cell's full text — sentence by sentence, not a keyword
   scan. Look for:
   - **誤字**: a wrong kanji/character that doesn't fit the intended word (a likely IME conversion
     slip, e.g. a homophone substituted for the correct word).
   - **脱字**: a missing character or word that makes the sentence ungrammatical or leaves an action
     without its object/verb.
   - **衍字**: an accidentally duplicated character or word (e.g. "確認するする").
   - **助詞の誤り**: a wrong particle that changes or breaks the meaning (e.g. "を" where "が" was
     needed, an action that reads backwards from what's clearly intended).
   - **Unmatched brackets/parentheses/quotes** within one cell.
   - **Garbled leftover fragments**: a word or clause that doesn't belong to the sentence it's sitting
     in — often the tell-tale sign of a copy-paste edit where the old text wasn't fully replaced.
3. Before flagging anything, run the struck-through/gray-color scan on that row/cell (per the shared
   doc) and drop it if marked deprecated.
4. Cross-check every candidate against context before reporting it as a real error:
   - Is this actually a deliberate project term or abbreviation, not a typo? (Check whether the same
     spelling appears consistently elsewhere in the project — if so, it's house style, not an error.)
   - Does a near-identical sentence exist elsewhere in the same doc (a sibling row in a repeating
     list, a parallel branch for a different condition)? Diffing against that twin is often the
     fastest way to confirm a real error versus a deliberate variation.
   - Could ambiguous phrasing be intentional shorthand common in this project's docs, rather than a
     mistake? If genuinely unsure, report it as a lower-confidence note rather than a confirmed error
     (see Reporting).

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on. Write in Japanese.

For each real finding: cite the exact sheet name and cell coordinate, quote the erroneous text
verbatim, state plainly what's wrong (誤字/脱字/衍字/助詞の誤り and which), and give the corrected
text when it's obvious. Order by impact: an error that changes or obscures the meaning of a processing
rule or condition belongs before a cosmetic misspelling that doesn't affect understanding. Distinguish
confirmed errors from "this reads oddly but might be intentional shorthand" judgment calls, and mark
the latter as needing the designer's confirmation rather than asserting it as a defect.

Omit entirely: a tally of how many cells/sections were proofread, "確認済み・問題なし" notes for
clean sections, and narration of your own process (which files you dumped, how you sampled). The one
exception worth a one-line mention: a whole prose section that couldn't be read at all (e.g. it was
entirely struck-through, or a sheet's used range exceeded the dump cap) — that's a real coverage gap
worth flagging, not a process detail.
