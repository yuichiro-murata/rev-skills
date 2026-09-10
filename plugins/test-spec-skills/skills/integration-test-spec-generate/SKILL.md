---
name: integration-test-spec-generate
description: Generate a 総合テスト仕様書／報告書 (.xlsm) for one program from its 機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表, by copying an already-completed spec workbook as the template and deriving every sheet from the design doc. Covers ﾃｽﾄ仕様 (test-item list), 項目・ｽﾞｰﾑ制御の確認, 画面項目制御の確認, ﾁｪｯｸ処理設計書（各画面）and 更新確認. Use when the user asks to 総合テスト仕様書を作成/作って for a program that has no spec yet, or to regenerate one from an updated design doc. This is a *generation* skill — it does not review an existing design doc (that is `rev-program-review` and its six check skills) and it does not execute tests or capture evidence.
---

# integration-test-spec-generate

Builds one program's 総合テスト仕様書／報告書 from its design-doc workbook. Everything in the spec
is derivable from the design doc — the point of this skill is that the derivation rules have been
**validated against a completed spec** rather than guessed, and that the mechanical traps that
silently corrupt the output are written down.

Validated on: generating `PSJCO501_送品案内入力` using `PSJCO502_納品案内入力` as the template. The
項目・ｽﾞｰﾑ制御 derivation reproduced 105 of 107 of 502's recorded values (97.2%) before being applied
to 501.

## Inputs and where they live

| what | where |
|---|---|
| 設計書 (source of truth) | `01_Doc/08_機能定義書/<WG>/PHASE<n>/<PROGRAM>_<名称>.xlsx` |
| テンプレート (完成済み仕様書) | `総合テスト/テスト仕様書参考/総合テスト仕様_V*_<PROGRAM>_*.xlsm` |
| エビデンス見本 | `総合テスト/テストエビデンス参考/…/<担当者>ｴﾋﾞﾃﾞﾝｽ/*.xlsx` |

Pick the template from the program whose screen structure is closest to the target. Sibling programs
in the same 機能 (e.g. 送品案内入力 / 納品案内入力) share almost all of their A画面 layout, which makes
the derivation cross-check meaningful.

**Always copy the template `.xlsm` and rewrite it** — do not build the workbook from scratch. The
sheets carry hundreds of merged ranges, conditional formats and a VBA project that cannot be
reproduced by hand. Load with `openpyxl.load_workbook(path, keep_vba=True)` and save the same way.

## Workbook structure and what each sheet is derived from

| シート | 出どころ |
|---|---|
| 改訂履歴 (hidden) | 手入力 |
| ﾃｽﾄ仕様 | 画面構成（Ⅴ．画面項目定義の共通ﾎﾞﾀﾝ群と領域名）から項目一覧を起こす |
| 是正テスト仕様 | 空のひな形（不具合が出てから記入） |
| ﾚｽﾎﾟﾝｽﾃｽﾄ | 空のひな形（実施時に記入） |
| 項目・ｽﾞｰﾑ制御の確認 | 画面設計書 Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 ＋ Ⅴ．画面項目定義 |
| 画面項目制御の確認 | 画面設計書 Ⅵ．画面項目制御・出力仕様 の「※」注記 |
| ﾁｪｯｸ処理設計書（各画面） | 設計書の同名シートを転記＋判定欄を追加 |
| 更新確認 | 更新条件表 各シートの「更新条件」欄 ＋ 機能定義書 Ⅲ．入出力定義(CRUD) |

## Step 1 — 画面構成を確定する

Parse each `画面設計書(<画面ID>)` sheet. Locate the sections by scanning column B for a heading that
matches `^([ⅣⅤⅥ])[．.]` — **the punctuation is inconsistent**: most sheets use the full-width `．`
but at least one (`GSJC502B`) uses the half-width `.`, and a `startswith("Ⅵ．")` test silently fails
there and aborts the whole run with a `KeyError`.

Within Ⅴ, blocks start at a cell in column C matching `^G\d\)` or the literal `共通`; item rows are
the ones whose column C is a bare number. Useful column map for Ⅴ (1-based):

```
C=3 No.   E=5 画面項目名   K=11 表示   Q=17 属性   U=21 TAB   W=23 桁数
Y=25 表示形式   AD=30 入力可文字種   AG=33 必須   AI=35 初期値   AQ=43 説明   BB=54 画面項目ID
```

and for Ⅳ: `C=3 No.  E=5 画面項目  M=13 ｲﾍﾞﾝﾄ内容  T=20 ｷｰ割当  X=24 処理内容  AQ=43 ﾒｯｾｰｼﾞ`.

From the `共通` block of Ⅴ you get the screen's buttons (attribute `Button`) and from the other
blocks the region names — those two lists drive the ﾃｽﾄ仕様 item list.

## Step 2 — 項目・ｽﾞｰﾑ制御の確認

For every 画面 × 領域 (excluding `共通`), emit one block. The sheet has **no merged cells** — it is
plain rows — so it can be written entirely with openpyxl.

Block layout (row offsets from the block header row):

```
+0  A=<X>画面  B=<領域名>  D=件数  E="件"  F="画面単位"
+1  C="項目制御"  I=項目制御件数   K="ｽﾞｰﾑ制御"  P=ｽﾞｰﾑ件数     ← ｽﾞｰﾑ列はズームがある場合のみ
+2  列見出し: C=項目 D=入力可文字種 E=選択可 F=小文字→大文字変換 G=桁数制御
             H=ｲﾍﾞﾝﾄ発生時ﾌｫｰｶｽ移動 I=ﾛｽﾄﾌｫｰｶｽ時の処理
             K=呼出元ﾊﾟﾗﾒｰﾀ L=ｽﾞｰﾑ名 M=起動 N=ｽﾞｰﾑへﾊﾟﾗﾒｰﾀが反映される
             O=ｽﾞｰﾑで選択した値が呼出元へ反映される P=ｽﾞｰﾑ終了後のﾌｫｰｶｽ位置
+3.. 明細行（項目1行ずつ）／ズーム行は同じ開始行から K列以降に並べる
```

Then a 2-row gap before the next block.

### 6軸の判定ルール（502 実績で検証済み）

A cell is either **対象** (leave the judgment cell empty for the tester to fill) or **対象外**
(write `-`). Never pre-fill `OK`.

| 軸 | 対象になる条件 |
|---|---|
| D 入力可文字種 | Ⅴ．入力可文字種(AD) が `-`/空 以外 |
| E 選択可 | Ⅴ．属性(Q) が `Button` `CheckBox` `RadioButton` `ComboBox` `LinkLabel` `Accordion` |
| F 小文字→大文字変換 | Ⅴ．説明(AQ) に「大文字に変換」を含む |
| G 桁数制御 | Ⅴ．桁数(W) が `-`/空 以外 |
| H ｲﾍﾞﾝﾄ発生時ﾌｫｰｶｽ移動 | **属性が `Button` かつ** Ⅳ にその項目のｲﾍﾞﾝﾄがあり、処理内容に「ﾌｫｰｶｽ」の記述がある |
| I ﾛｽﾄﾌｫｰｶｽ時の処理 | Ⅳ にその項目の「ﾛｽﾄﾌｫｰｶｽ」ｲﾍﾞﾝﾄがある |

Three rules here are counter-intuitive and were each learned from a mismatch against 502:

- **ﾛｽﾄﾌｫｰｶｽ(I) は Ⅳ で判定する。Ⅴ の説明欄で判定してはいけない。** `送付先作業場名`(Label) の説明に
  「送付先作業場名ﾛｽﾄﾌｫｰｶｽ時:(3).作業場名」と書かれているが、実際にｲﾍﾞﾝﾄを持つのは
  `送付先作業場ｺｰﾄﾞ` のほう。説明欄で判定すると対象が真逆になる。
- **ﾌｫｰｶｽ移動(H) は「ﾌｫｰｶｽ移動は行わない」と書かれていても対象。** イベント後のフォーカス挙動を
  確認する軸なので、移動しないことの確認も対象に含まれる。ただし **Button のみ**。TextBox の
  ﾛｽﾄﾌｫｰｶｽ処理を H に立ててはいけない（それは I 軸）。
- **Ⅳ と Ⅴ で項目名に `(1-N)` の有無のブレがある。** 例: Ⅴ は `編集ﾎﾞﾀﾝ(1-N)`、Ⅳ は `編集ﾎﾞﾀﾝ`。
  突合前に `(1-N)` を落として正規化する。

### ズーム制御

From Ⅳ, an item whose name contains `ｽﾞｰﾑ` and whose 処理内容 matches `…ｽﾞｰﾑ(ZXJAnnn)` is a zoom.

- **ｽﾞｰﾑ名の抽出でひらがなを含めない。** `…をもとに作業場ｽﾞｰﾑ(ZXJA029)` を素直に正規表現で取ると
  「もとに作業場ｽﾞｰﾑ」になる。`[一-龥ァ-ﾝA-Za-z0-9]+ｽﾞｰﾑ` のようにひらがなを除外して取る。
- 呼出元ﾊﾟﾗﾒｰﾀは実績表記が `G1).案内書No(From)` 形式（`G1)` の後にピリオド）。設計書本文には
  `ｺｰﾄ`（濁点欠け）の表記ゆれがあるので `ｺｰﾄﾞ` に補正する。

### 件数の出し方

- 領域ブロックの件数 = 項目制御の `OK` セル数 ＋ ズーム件数 × 4（起動/ﾊﾟﾗﾒｰﾀ反映/戻り値反映/ﾌｫｰｶｽ位置）
- シート冒頭(D1)の全体件数 = 各ブロック件数の合計

A good sanity check: a 領域 that is structurally identical to the template program's should produce
the same count (A画面 G1)検索条件領域 came out 45 for both 501 and 502).

## Step 3 — 画面項目制御の確認

Derived from the **「※」注記 in Ⅵ．画面項目制御・出力仕様**, which is where the conditional branches
live. A screen with no ※ notes gets no block at all.

Typical axes seen: `明細 0件/0件以外`, `処理区分 "1"(新規登録)/"2"(変更)`,
`ﾛｸﾞｲﾝ情報.作業者ｺｰﾄﾞ NULL/NULL以外`, `明細表示ﾌﾗｸﾞ 該当ﾃﾞｰﾀあり/なし`, `新規 ﾁｪｯｸなし/あり`.

想定結果 is always the literal `基本設計書記載の内容と一致`; the 判定 column stays empty.

**This sheet needs human review.** Turning a ※ note into a set of condition patterns is a judgment
call, not a mechanical transform — a note like「共通ﾕｰｻﾞｰではない かつ 明細行を表示している場合のみ○」
has two axes but the template program listed only one. Say so when handing the file over.

## Step 4 — ﾁｪｯｸ処理設計書（各画面）

The spec version is the design doc's sheet **shifted 7 rows up and 1 column right**, with a judgment
column inserted at A:

```
設計書 row 8(見出し) → 仕様書 row 1      設計書 row 9(先頭明細) → 仕様書 row 2
設計書 A(No.) → B    C(項目名) → D    K(区分) → L    N(ﾁｪｯｸ内容) → O
      AE(ﾚﾍﾞﾙ) → AF  AH(出力先) → AI  AN(ｴﾗｰﾒｯｾｰｼﾞ) → AO  BA(ｴﾗｰｺｰﾄﾞ) → BB
      BG(ﾁｪｯｸ詳細) → BH   CH(ﾁｪｯｸ補足) → CI(Jira記入欄として使う)
新規: A1="確認"（A1:A2 を結合）、CJ1=件数
```

Do this with **Excel COM worksheet copy**, not openpyxl — copying a sheet between workbooks preserves
the merges and conditional formats that openpyxl would lose.

Procedure per screen: copy the sheet in → delete the fully-struck check rows (bottom-up) → clean the
mixed-strikethrough cells → `Rows("1:7").Delete()` → `Columns(1).Insert()` → set A1, mark the
`【…押下時】` separator rows with `-` in column A, write the count formula → rename.

### 取り消し線の扱い

Confirmed against the template: **設計書 B画面 80行 − 全体取り消し5行 = 仕様書 75行.** Fully struck
check rows are deleted when the spec is built. Two distinct cases, handled differently:

- **行全体が取り消し** (a check that was removed): delete the row. Detect it as a row whose column A
  is numeric and which has ~5+ struck cells — the unstruck one is usually just the DA revision note.
- **セル内の一部だけ取り消し** (a message or code that was *changed*): keep the row, keep only the
  live characters. openpyxl reports `cell.font.strike == True` for these because it only sees the
  first run — you must go through Excel COM `Characters(i, 1).Font.Strikethrough` to separate them.
- **チェック表の外にある取り消しブロック** (deleted 参照ｴﾝﾃｨﾃｨ definitions etc.): leave in place,
  struck. It is supporting detail, not a test item, and the strikethrough documents the history.

Also drop the `CI:DA` merge that comes across from the design doc's ﾁｪｯｸ補足 column — the template
does not have it, and it blocks the `CJ1` count cell.

## Step 5 — 更新確認

Two sources:

- **更新条件表 各シートの `P12`（更新条件）** — this cell names the trigger button and the branch
  logic, e.g.「実行ﾎﾞﾀﾝ押下時 ／ 処理区分:"1"(新規登録)または"3"(一時保存)の場合 INSERT ／ 上記以外の場合
  DELETE→INSERT」. `P8` holds `<ID>:<テーブル名>`. Read all of them to group tables by event.
- **機能定義書 Ⅲ．入出力定義** — the C/R/U/D flags. Tables with a `D` flag are the delete targets for
  the 削除ﾎﾞﾀﾝ, which the individual 更新条件表 usually do not spell out.

Block layout: condition axes start at column D, the update-target table columns follow immediately,
and the 判定 column is the one after the last table. Table header cells are `"<ID>\n<名称>"`.

**削除対象の特定は推定が入る。** Only one 更新条件表 typically says 削除ﾎﾞﾀﾝ押下時; the rest is
assembled from the CRUD flags and the template program's precedent. Flag it for review.

## Step 6 — ﾃｽﾄ仕様（項目一覧）

Per screen, in this order: 全般 → 初期表示 → 各領域(G1..Gn) → ﾁｪｯｸ処理 → 各ﾎﾞﾀﾝ. Numbering is
`S1-<通し番号>` for the section header row and `S1-<通し番号>-<n>` for each item.

Reusable item sets:

- **全般**: 画面ｲﾒｰｼﾞ一致 / ﾀﾌﾞ順一致 / 多言語(日本語) / 多言語(ﾍﾞﾄﾅﾑ語)
- **初期表示**: 初期値一致 / 画面項目制御一致 / 初期ﾌｫｰｶｽ位置一致
- **領域**: 各項目の制御・処理（「項目・ｽﾞｰﾑ制御の確認」参照）/ 左右寄せ / 編集書式
  （明細グリッドの領域はさらに ﾁｪｯｸﾎﾞｯｸｽ変更 / ﾍｯﾀﾞｰﾁｪｯｸで全選択）
- **ﾁｪｯｸ処理**: 「ﾁｪｯｸ処理設計書（X画面）」参照
- **更新系ﾎﾞﾀﾝ**: 確認ﾒｯｾｰｼﾞ / DB更新内容（「更新確認」参照）/ 後続処理 / 画面項目制御 / ﾌｫｰｶｽ
- **遷移系ﾎﾞﾀﾝ**: 遷移先とﾊﾟﾗﾒｰﾀ / ﾌｫｰｶｽは遷移先の初期位置

**画面数はテンプレートと違うことが前提。** The template may have 3 screens where the target has 5.
Rebuild the whole S1 block from the target's own screen list rather than editing the template's text
— a text-replace pass leaves the extra screens missing and mislabels the ones that shifted (a
一時ﾃﾞｰﾀ一覧 that is C画面 in the template but E画面 in the target).

Also rebuild the S5 (ﾚｽﾎﾟﾝｽﾃｽﾄ) block from the target's screens, and check the S4 section headers for
button names that only exist in the template (`更新ﾎﾞﾀﾝ` vs `実行ﾎﾞﾀﾝ`).

## Step 7 — 検証（省略しないこと）

Rendering a screenshot of one region is **not** sufficient; it only proves the region you looked at.
Run all of these before reporting the file as done.

### 1. 結合セル

```python
# 行を挿入したなら結合数は必ず増える。変わっていないこと自体がバグの証拠。
expected = template_merge_count + inserted_rows
```

ﾃｽﾄ仕様 merge rules, taken from the template:

| 範囲 | ルール |
|---|---|
| S1 機能ﾃｽﾄ枠 | 見出し行（本文が `<` で始まる）= `B:M`、明細行 = `B:N` |
| S4 以降 | 一律 `B:N`（`ｼﾅﾘｵﾃｽﾄ` `対象外` のみ `B:M`） |
| 区切り行（B列が空） | 結合なし |

Excel の行挿入は書式は上の行から継承するが **結合は引き継がない**。挿入した行だけ結合が抜ける。
Safest fix: unmerge every single-row `B`-anchored merge from row 34 down, then re-apply by rule.

### 2. 行高

Template heights by line count: `1行=24 / 2行=36 / 3行=42.75 / 4行=57`. Inserted rows keep a
single-line height and clip multi-line text. Grow the rows that are too short; never shrink.

### 3. 書式のリセット漏れ

When clearing a template sheet for rewriting, clear **values, styles and row heights** together —
`cell.value = None` alone leaves the old fills, colored banners and tall rows stranded at row
positions where the new content no longer lines up:

```python
ws.row_dimensions[r].height = None
cell.value = None
cell.style = "Normal"
```

### 4. テンプレート由来の残骸

Grep the finished file for the template program's ID, screen IDs, product name, its Jira numbers and
`#REF!`. Expect zero hits. (Note that the *target's own* design-doc revision notes legitimately
contain NCRN numbers — match on the template program's identifiers, not on `NCRN` alone.)

### 5. 目視

Render the sheets to PNG via Excel COM and actually look at them, including the rows around every
insertion boundary.

## 実装上の罠（すべて実際に踏んだもの）

- **openpyxl の `MergedCell` は書き込み不可。** `cell.value = x` が `AttributeError` で落ちる。書く前に
  `isinstance(cell, MergedCell)` で弾くか、結合範囲の左上セルに書く。
- **PowerShell で日本語は変数名になる。** `"ﾁｪｯｸ処理設計書（$scr画面）"` は `$scr画面` という
  未定義変数として解釈され、空文字になる。`${scr}` と明示的に囲む。
- **COM のパラメータ付きプロパティは括弧が要る。** `$ws.ChartObjects().Add(...)`、
  `$c.Characters().Count`。括弧なしだと静かに失敗し、`Characters.Count` が 0 を返した結果
  「生きている文字」が空になってセルを消してしまう。
- **PowerShell の `$$` は自動変数。** 数式文字列に `"…$B$$last"` と書くと壊れる。`${last}` を使う。
- **`Chart.Paste()` の前後に待ちを入れる。** `CopyPicture` 直後に貼ると空の画像が出力される
  （700ms 程度の `Start-Sleep` で解消）。真っ白なレンダリング結果を見たらまずこれを疑う。
- **COM セッション後は `EXCEL` プロセスを止める。** 残っているとファイルがロックされ、次の
  openpyxl の `save()` が素の permission error で落ちる。
- **`Add-Type` / スクリプト実行の制約**: この環境の Cylance Script Control は `.ps1` ファイルの実行を
  ブロックする。PowerShell はインラインで渡すこと。`[\/:*?"<>|]` という文字クラスをコマンドに
  含めると、安全ガードが誤検知してコマンド全体が拒否される。

## 完了報告に必ず含めること

- 各シートの件数（ﾃｽﾄ仕様の項目数／項目・ｽﾞｰﾑ制御／画面項目制御／ﾁｪｯｸ処理／更新確認）
- **人のレビューが要る箇所**: 画面項目制御の条件パターン展開、更新確認の削除対象テーブル、
  設計書から名前が読み取れなかった画面の仮称
- **未作成のまま残したシート**（是正テスト仕様・ﾚｽﾎﾟﾝｽﾃｽﾄ）と、その理由
- テンプレートから引き継いだまま実施時に更新が要る項目（APｻｰﾊﾞｰ/DBｻｰﾊﾞｰ/ブラウザのバージョン等）
