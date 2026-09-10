---
name: rev-program-review
description: Entry point for a full REV of ONE program's design-doc workbook (機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表). Instead of unconditionally running every check, it first presents the 6 single-program check skills as a checkbox list (AskUserQuestion, multiSelect) so the user picks which checks to run BEFORE any dumping or analysis starts, then runs only the selected ones as one combined pass and reports their findings as a single merged report. Use whenever the user asks to REV/レビュー a design-doc workbook without naming the specific checks they want — in that case do NOT launch the individual check skills directly. Skip the checkbox prompt only when the user already named the checks, or explicitly asked for 全部/all/フルREV.
---

# rev-program-review

Runs a REV of one program's design-doc workbook with a **user-selected scope**. The 6 checks below
are individually expensive (each dumps and re-reads a large workbook, several also read WG-folder
DB-layout files and the 01_Doc common-design workbooks), so running all of them when the user only
wanted two wastes a lot of time and tokens. Ask first, then run only what was selected.

## The 6 single-program check skills

| # | skill | what it checks |
|---|-------|----------------|
| 1 | `design-doc-internal-consistency` | 内部相互参照（ｲﾍﾞﾝﾄ⇔処理概要、ﾚｽﾎﾟﾝｽ/画面項目ID/ﾁｪｯｸID登録、画面3点一致、排他制御） |
| 2 | `design-doc-io-table-check` | Ⅲ．入出力定義（CRUD一覧）の双方向網羅性 — 最重量・最高収穫のチェック |
| 3 | `xlsx-db-column-check` | 参照ｶﾗﾑがﾃｰﾌﾞﾙﾚｲｱｳﾄに実在するか＋ｺｰﾄﾞ/名称の二重保持ｱﾝﾁﾊﾟﾀｰﾝ(名称がLabelの場合のみ指摘) |
| 4 | `naming-standard-compliance` | 各ID採番規則・設計書記述ﾙｰﾙ準拠 |
| 5 | `design-doc-formatting-consistency` | ﾌｫﾝﾄｻｲｽﾞ/ｾﾙ結合のブレ（体裁の衛生） |
| 6 | `design-doc-typo-check` | 日本語の誤字脱字・変換ミス |

WG-folder-scoped comparison skills are **not** part of this menu — bundle those in only when the
user's request is itself folder-scoped, or they explicitly ask for them.

## Step 1 — 対象ワークブックの確定

Resolve which workbook is being reviewed (a program ID, a path, or the IDE selection). If the target
is genuinely ambiguous, ask for it in the same `AskUserQuestion` call as the scope question below —
don't burn a separate round trip.

## Step 2 — チェック項目の選択（チェックボックス）

Call `AskUserQuestion` with **two `multiSelect: true` questions** — the tool allows at most 4 options
per question, so the 6 checks are split 4 + 2. Present them exactly like this (labels in Japanese,
since the reviewers work in Japanese) — keep this grouping and this option order:

- Question 1 — header `整合性`, question「実施するチェックを選択してください（複数選択可）」
  - 「Ⅲ．入出力定義（CRUD）網羅チェック (推奨)」— `design-doc-io-table-check`
  - 「設計書内部の相互参照チェック (推奨)」— `design-doc-internal-consistency`
  - 「DBカラム実在チェック」— `xlsx-db-column-check`
  - 「ID採番・記述ルール準拠チェック」— `naming-standard-compliance`
- Question 2 — header `誤字・体裁`, question「実施する誤字・体裁チェックを選択してください（複数選択可）」
  - 「誤字脱字チェック」— `design-doc-typo-check`
  - 「フォントサイズ・セル結合のブレチェック」— `design-doc-formatting-consistency`

Both questions go in **one** `AskUserQuestion` call, so all 6 checkboxes appear in a single prompt.

Put the skill name in each option's `description` alongside a one-line summary of what it finds, so
the user can tell the options apart without knowing the skill names by heart. The user can select
none in a group — that group's checks are simply skipped.

### 「Other」の扱い

"Other" lets the user type a scope in free text. Two cases, and they mean opposite things:

- **Other with text** — honor exactly what they typed for that group, even if it names a check from
  the other group or narrows the scope further ("入出力定義だけ", "画面項目の順序だけ見て").
- **Other selected but left empty** — read it as "この観点は実施しない": skip **every** check in
  that group, list them under 未実施 in the report, and don't ask again. It is the deliberate way to
  opt a whole group out, so treat it exactly like selecting nothing in that group — never as a
  prompt to re-ask, and never as a reason to fall back to running the group's checks.

If **both** groups end up with no check to run (nothing selected, or empty "Other"), there is
nothing to review: stop, state plainly that no check was run and that the workbook was not dumped,
and do not fall back to running all 6.

### プロンプトを省略してよいケース

Skip Step 2 and go straight to Step 3 when:

- The user already named the checks ("入出力定義と誤字だけ見て", "誤字チェックして") — run exactly those.
- The user explicitly asked for everything ("全部", "フルREV", "6つ全部", "all") — run all 6.
- The user invoked one check skill directly by name — that skill runs standalone; this skill isn't involved.
- `AskUserQuestion` is unavailable (non-interactive / batch / subagent context) — fall back to all 6
  and **state in the report** that the full set was run because the scope couldn't be asked.

Re-ask the scope for each **new** REV request; a selection made for one workbook does not carry over
to the next one unless the user says "同じ観点で" or similar.

## Step 3 — 選択されたチェックの実行

1. **Dump the workbook once, up front** — before launching anything — and hand every check the
   resulting scratchpad text files instead of letting each one re-dump the same workbook. See
   `_shared/xlsx-excel-com-dump.md`, section "dump once, share the text" — that file ships **inside
   this plugin** (`<plugin root>/skills/_shared/`), not under `~/.claude/skills/`; glob
   `**/rev-skills/**/skills/_shared/xlsx-excel-com-dump.md` if the path doesn't resolve. (Exception:
   `design-doc-formatting-consistency` still needs live Excel COM access for its font/merge scan;
   only the bulk text dump is shared.)
   `**/rev-skills/**/skills/_shared/xlsx-excel-com-dump.md` typically matches several copies — the
   git-tracked marketplace copy (`.claude/plugins/marketplaces/rev-skills/plugins/rev-skills/...`)
   **and** one or more older snapshots under `.claude/plugins/cache/rev-skills/<version>/...`. Always
   read the **marketplace** copy: the cache lags behind it, and a stale cached copy has already cost
   a run — the `PSJCO309` dump failed on the `Add-Type` CS0675 bitwise-or error that the marketplace
   copy documents a fix for but the cached `1.1.1` copy predates. The cache is keyed on the
   `version` in `.claude-plugin/plugin.json`, so editing a skill without bumping that version leaves
   every cached copy stale **forever** — the loader sees a version it already has and never re-copies.
   Bump the version in the same commit as any skill-content change.
2. **Build the reference-master index in the same pre-launch step, if a selected check needs it** —
   see `_shared/reference-index.md`. Today only `design-doc-internal-consistency` does (the
   画面項目辞書 index). That doc instructs the *agent* to stop rather than build it, so skipping this
   step silently drops that check's dictionary-registration test. Build **one index per dictionary
   file the program's screen-item IDs route to** — that doc's routing table maps each ID prefix to
   its `82.画面項目辞書_*.xlsx`, and a program using shared `XJZ`/`SJZ` items needs the `_共通` file
   as well as its own WG's. Subset each to the program's IDs, and pass both paths per file in the
   agent's prompt — the subset to read, the full index to grep.
3. Run the selected checks as one combined pass — in parallel background agents when there are
   several. Follow each selected skill's own SKILL.md as the authority for how that check is done;
   this skill only decides *which* checks run.
4. Do **not** post a status update as each agent finishes. Wait until every check in the batch has
   completed, then compose and post **one** merged report.

## Step 4 — 報告

One combined report, findings grouped by check, with the selected scope stated at the top so the
reader knows what was and wasn't looked at — e.g.:

```
## REV結果: PXJCO201_処置指示登録
実施チェック: Ⅲ．入出力定義(CRUD)網羅 / 設計書内部の相互参照 / 誤字脱字
未実施: DBカラム実在 / ID採番・記述ルール / フォントサイズ・セル結合
```

Never imply the workbook passed checks that weren't run.
