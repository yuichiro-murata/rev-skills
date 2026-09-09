# Reference-master indexes: read a lookup table, not a whole workbook

`xlsx-excel-com-dump.md` covers the workbook **under review**. This file covers a reference master a
REV skill consults as ground truth. Those are large, they change rarely, and a check that only needs
to look up an ID needs two or three columns out of forty-odd — so for that shape of check, build a
small index once and let the agents read the index.

## What this file covers, and what it does NOT

| Reference master | How to read it |
|---|---|
| `82.画面項目辞書_<WG>.xlsx` | **Index it — builder below.** |
| `<TableID>_<name>.xlsx` (ﾃｰﾌﾞﾙﾚｲｱｳﾄ) | **Not yet indexed** — keep using the `xlsx-dumps` cache. |
| `06-06_DB一覧_<WG>.xlsx` | **Not yet indexed** — keep using the `xlsx-dumps` cache. |
| `06-03_帳票一覧_<WG>.xlsm` | **Not yet indexed** — keep using the `xlsx-dumps` cache. |

**Do not generalise the pattern to the other three yourself.** Each was tried and each hid a
structural trap that a hand-rolled index gets wrong silently:

- `06-06_DB一覧_<WG>.xlsx` carries **two overlapping table lists on two visible sheets** with
  different header labels — `作成状況一覧` (row 1, `ﾃｰﾌﾞﾙID`/`ﾃｰﾌﾞﾙ名`, 353 ids) and `DB一覧`
  (row 35, `ID`/`名称`, 219 ids). Indexing only the first drops 36 tables and reports them as
  unregistered.
- A table-layout index has **one source file per table**, so the single-source freshness contract
  below does not apply to it and there is no agreed replacement yet.
- The table IDs a program cites include views and suffixed forms (`VXJCM004_31`, `TXJAM061WF`,
  `VXJCM004_31_ALL`); an index built from a `T`-only ID pattern silently omits them.

Only `design-doc-internal-consistency` reads this file today — it is the only check that needs
`id → 画面項目名`. `xlsx-db-column-check` and `design-doc-io-table-check` keep using the dump cache
until their masters get builders of their own.

## Who builds the index: the orchestrator, once, before launching agents

The index is a shared artefact, exactly like the workbook dump. Build it in the orchestrating
session **before** any agent starts, and hand agents the path.

**Never tell parallel agents to "build or reuse" it.** Several sibling skills can want the same
master in one REV; if two build it at once they write the same file concurrently and one of them
reads a torn index. Same reasoning as "dump once, share the text" in `xlsx-excel-com-dump.md`.

If you are a sub-agent and the index you were pointed at is missing or stale, **say so and stop** —
do not build it yourself, and do not fall back to opening Excel unless your prompt allowed it.

## Which dictionary file: route by the ID's own prefix, not by the reviewing WG

`82.画面項目辞書_*.xlsx` is split by JOBコード, not by who is doing the review. A 工程管理 program that
puts a shared item on its screen cites an `XJZ…`/`SJZ…` ID, and that ID is registered in
`82.画面項目辞書_共通.xlsx` — it is not in `_工程管理` at all. Look for it there and you report a
registered item as unregistered.

| ID prefix | Dictionary file (under `01_Doc/04_共通設計/`) |
|---|---|
| `XJZ` / `SJZ` | `82.画面項目辞書_共通.xlsx` |
| `XJA` / `SJA` | `82.画面項目辞書_基準情報.xlsx` |
| `XJB` / `SJB` | `82.画面項目辞書_受注出荷.xlsx` |
| `XJC` / `SJC` | `82.画面項目辞書_工程管理.xlsx` |
| `XJD` / `SJD` | `82.画面項目辞書_品質管理.xlsx` |

A `MENU`-prefixed ID carries its routing prefix immediately after `MENU` — `MENUXJCP00` → 工程管理,
`MENUXJZP00` → 共通.

This is the routing rule `04.ﾒｯｾｰｼﾞ管理_*.xlsx` already follows (see `xlsx-excel-com-dump.md`): a
foreign prefix routes to that other WG's own file, **never** to a same-named sheet embedded in the
copy you happen to have open. In this file that matters twice over, because the sheets the 品質管理
and 受注出荷 copies carry under other WGs' names are not dictionaries at all — see the builder's
sheet table below.

**Build one index per dictionary file the program's IDs route to — never one merged index.** The
freshness contract below is single-source: a merged index has no meaningful `# source-mtime-utc`.
Collect the distinct prefixes from the program's dump, build an index for each file they name,
subset each, and hand the agent all of them. Most programs touch one or two files. The builder needs
no per-file configuration — it locates dictionary sheets by their header, so the same call handles
every one of the five.

## Measured effect

Measured on `82.画面項目辞書_工程管理.xlsx` at source length 1,945,616 / mtime
`2026-09-09T09:30:29Z` — the masters are edited during a review cycle, so treat these as an order of
magnitude and re-measure if the numbers matter. The comparison is against the dump this check used
to take: the `XJC(*`/`SJC(*` sheets only, which is all any check ever read from this file (`改訂履歴`,
`Sheet1`, `翻訳リスト` are never used by any check; the whole workbook dumps to 4,771,257 chars):

| Artefact | Chars | Est. tokens |
|---|---:|---:|
| Scoped 2-sheet dump (the previous approach) | 4,215,115 | ~1,318,000 |
| Index of the same two sheets | 56,333 | ~28,200 |
| Index subset to the 106 IDs one program cites | 2,227 | ~1,100 |

**46.7x smaller by tokens, and ~1,150x once subset.** Token estimates use ASCII÷3.5, half-width
katakana×1.0, full-width×0.9, other×1.0 — the same method throughout this repo; character counts are
exact. Note the character ratio (74.8x) and the token ratio (46.7x) differ: the dump is mostly ASCII
coordinate tokens, which are ~3.5 chars each, while the index is mostly Japanese names.

## Format: sectioned CSV

One record per line, hierarchy carried by `#` comment lines rather than repeated columns:

```csv
# index: screen-item-dictionary
# source: 82.画面項目辞書_工程管理.xlsx
# source-mtime-utc: 2026-09-09T08:08:21.9838183Z
# source-length: 1945178
# columns: id,name
# sheet: XJC(工程)
XJC8046,管理No
MENUXJCP00,工程管理
# sheet: SJC(工程)再開発追加分 
SJC8613,【対象不良内容(履歴)】
SJC8614,YOTO
```

**Why this shape and not a Markdown table.** Measured on the 2,215-record 基準情報 dictionary,
carrying identical information: MD table 24,888 tokens, flat CSV repeating sheet/group on every row
38,867 (+56%), **sectioned CSV 20,894 (-16%)**. Flat CSV is the worst of the three — do not
"simplify" the section headers into columns.

The bigger reason is correctness. A Markdown table cannot represent a value containing `|` or a
newline, and its ` | value | ` convention silently strips leading and trailing whitespace. In the
工程管理 dictionary's ID and name columns alone there are **2 values containing a newline and 18 with
significant leading/trailing whitespace** (`'品目ﾏｽﾀ無し及び特殊作業ﾘｽﾄ '`, `'  /  /  '`,
`'  月  日'`). Those trim to something else, and a name-match check then reports a mismatch that does
not exist. CSV with proper quoting is lossless for all of them.

**Quoting rule — RFC4180 plus one extension.** Quote a value when it contains a comma, a double
quote, CR or LF, **or when it has leading or trailing whitespace**. The last clause is not in
RFC4180 and is not optional here: most parsers strip unquoted surrounding spaces, which is precisely
the data loss above. Escape an embedded double quote by doubling it.

```powershell
function CsvQ($s) {
    if ($null -eq $s) { return "" }
    $t = [string]$s
    if ($t -match '[",\r\n]' -or $t -ne $t.Trim()) { return '"' + $t.Replace('"','""') + '"' }
    return $t
}
```

### Reading the index back: one record is not always one line

Because the format is lossless, a value containing a newline makes **one record span several
physical lines**. In the 工程管理 dictionary index that is 2 records (`SJC7009`, `SJC0633`) spread
over 6 extra physical lines, and one more record has a comma inside its quoted name
(`SJC0600,"0001～ZZZZ (0～9,A～Z(I,O,Q除く))"`).

So a line-oriented read of the index is wrong in three specific ways, all silent:

- **Counting records by counting lines over-counts.** The 工程管理 index has 3,288 records on 3,294
  data lines. (Confirmed the hard way: the discrepancy looked like a builder bug and cost a round of
  debugging.)
- **`$_ -split ','` splits inside quotes**, so the `name` field comes back truncated for the record
  above.
- **A continuation line matches no id and does not start with `#`**, so a naive filter drops it and
  leaves the record's first line with an unterminated quote — a malformed CSV handed to an agent.

Grep/`Select-String` for a single id is fine (an id never contains a newline, and it is the first
field). Anything that walks records must track quote state — see the subsetting snippet below.

## Struck-through and grayed-out entries

A retired dictionary entry must not count as registered, or the check reports "this ID exists" about
something the design has withdrawn. Apply the same cascade as the workbook dump: ask the whole
`UsedRange` once and only walk cells when it comes back mixed.

**A partially-struck cell (`Font.Strikethrough` returns `DBNull`) is NOT a dead entry.** In this
project it is overwhelmingly an entry **renamed in place** — the old name struck through, the new
one typed next to or below it in the same cell. Treating `DBNull` as dead drops the whole row, and
the check then reports a live, registered entry as unregistered.

That is not hypothetical. The rule was written the other way first, and on one program's REV it
produced a false "帳票一覧未登録" finding for two of that program's own reports (`RSJC033`,
`RSJC035`) — the registry rows existed, they had merely been renamed in place — and three separate
agents repeated it. In this dictionary the same shape occurs on `SJCRSJC006`
(raw `現品票識別ｶｰﾄﾞ(共通)`, live `識別ｶｰﾄﾞ(共通)`).

Reconstruct the live remainder character by character, exactly as the workbook dump does, and drop
the row only when the **key** cell has no live text left:

```powershell
function IsGray([double]$argb) {
    $r = [int]$argb -band 0xFF; $g = ([int]$argb -shr 8) -band 0xFF; $b = ([int]$argb -shr 16) -band 0xFF
    return (($r -eq $g) -and ($g -eq $b) -and ($r -gt 80) -and ($r -lt 220))
}
# $null = the cell is dead (fully struck, gray, or nothing left once struck runs are removed).
function LiveText($cell) {
    $raw = "$($cell.Value2)"
    $col = $cell.Font.Color
    if (-not ($col -is [System.DBNull])) { if (IsGray $col) { return $null } }
    $s = $cell.Font.Strikethrough
    if ($s -is [System.DBNull]) {
        $lv = ""
        for ($i = 1; $i -le $raw.Length; $i++) {
            $ch = $cell.Characters($i, 1)
            if (-not $ch.Font.Strikethrough) { $lv += $ch.Text }
        }
        if ($lv.Trim().Length -eq 0) { return $null }
        # Removing the struck old value can leave the line break that separated it from the new
        # one, so the result starts with "\n". Strip leading/trailing CR/LF from a RECONSTRUCTED
        # value only. Never strip spaces — see the significant-whitespace values above.
        return ($lv -replace '^[\r\n]+','' -replace '[\r\n]+$','')
    }
    if ($s -eq $true) { return $null }
    return $raw
}
```

Measured on this file: 1 entry retired outright and 1 renamed in place, out of 3,288.

## `Value2` is UsedRange-relative; `Cells.Item` is sheet-absolute

`$used.Value2` is indexed from 1 **within the UsedRange**, but `$sheet.Cells.Item(r,c)` is absolute
on the sheet. Mixing them reads a *different cell* on any sheet whose UsedRange does not start at
`A1`, with no error.

This bites here specifically: in `82.画面項目辞書_工程管理.xlsx` the **`XJC(工程)` sheet's UsedRange
starts at column B** (`SJC(工程)再開発追加分 ` starts at A). A builder that takes values from the
array and asks `Cells.Item` about formatting was testing strikethrough one column to the left for
every one of that sheet's 1,460 rows.

Compute the origin offset once per sheet and apply it to every absolute access:

```powershell
$r0 = $used.Row - 1
$c0 = $used.Column - 1
$cell = $sheet.Cells.Item($r + $r0, $c + $c0)     # $r,$c are array (UsedRange-relative) indices
```

For the same reason, **column positions resolved from the array are array-relative**. In this file
the 画面項目ID group sits at array column 3 on `XJC(工程)` and array column 4 on
`SJC(工程)再開発追加分 ` — which is the *same* absolute column 4 on both, seen through different
offsets. Resolve by label, use the label's array index for values, and add the offset for formatting.

## Freshness: rebuild when the source file changes

The index carries `# source-mtime-utc` and `# source-length` from the source workbook. Before using
it, stat the source and compare both; if either differs, or the index does not exist, rebuild. This
is the same contract as the `xlsx-dumps` cache, and it matters more here — these masters are edited
during a review cycle, and a stale dictionary produces "unregistered ID" findings for IDs somebody
added yesterday. Never reuse an index without the check.

```powershell
function Test-IndexFresh($indexPath, $sourcePath) {
    if (-not (Test-Path -LiteralPath $indexPath)) { return $false }
    $src  = Get-Item -LiteralPath $sourcePath
    $head = Get-Content -LiteralPath $indexPath -TotalCount 8 -Encoding UTF8
    $m = ($head | Where-Object { $_ -like '# source-mtime-utc: *' }) -replace '^# source-mtime-utc: ',''
    $l = ($head | Where-Object { $_ -like '# source-length: *' })    -replace '^# source-length: ',''
    return ($m -eq $src.LastWriteTimeUtc.ToString('o')) -and ($l -eq [string]$src.Length)
}
```

**This works only for an index built from ONE source file**, where both header lines sit in the
first 8 lines. An index aggregating many sources needs a different contract — which is one of the
reasons the table-layout index is not in this file yet.

Cache root: `<user home>\.claude\skills\_cache\reference-index\`, one `.csv` per source file, named
`idx_<kind>_<source file stem>.csv` — e.g.
`idx_screen-item-dictionary_82.画面項目辞書_工程管理.csv`.

## Builder: screen-item dictionary (`82.画面項目辞書_<WG>.xlsx`)

Three things about this file's structure, each of which produces a plausible-looking but useless
index if you get it wrong:

**The ID is two columns concatenated.** Row 6 carries the group headers (`画面項目ID`, `画面項目名`),
row 7 the sub-headers (`ＩＤ`, `連番`, `日本語`, `ベトナム語`). The 画面項目ID group spans the `ＩＤ`
column (a category: `MENU`, `XJC`, `SJC`) and the `連番` column (the rest: `8046`, `GXJC101A`,
`XJCP00`). What a design doc cites is the two joined: `XJC` + `8046` = `XJC8046`,
`MENU` + `XJCP00` = `MENUXJCP00`. Reading either column alone gives you IDs that appear in no design
doc at all — confirmed by building it wrong first, which produced a 3,349-row index full of
`MENU,工程管理` rows and no `XJC8046` anywhere.

**Resolve the columns by label, never by a fixed position** — and read the offset section above,
because the two sheets' array positions differ while their absolute positions do not.

**Find the dictionary sheets by their header, not by their name.** Every real dictionary sheet, in
every WG's copy, has the identical layout — `画面項目ID` / `画面項目名` on row 6 at absolute columns 4
and 6, and `ＩＤ` / `連番` / `日本語` / `ベトナム語` on row 7 at absolute columns 4-7. Confirmed on all
five files, so **one builder handles all of them and no per-file dump method is needed**:

| Dictionary | Dictionary sheets | Visible sheets that are NOT dictionaries |
|---|---|---|
| `_共通` | `XJZ(共通_JAGUR)`, `SJZ(共通_刷新)` | `改訂履歴`, `Sheet1` |
| `_基準情報` | `XJA(基準_JAGUR)`, `SJA(基準_刷新)` | `改訂履歴` |
| `_工程管理` | `XJC(工程)`, `SJC(工程)再開発追加分 ` | `改訂履歴`, `Sheet1`, `翻訳リスト` |
| `_品質管理` | `XJD(品質_JAGUR)`, `SJD(品質_刷新)` | `XJZ(共通)`, `XJA(基準)`, `XJB(受注)`, `XJC(工程)`, `XJE(生産）` (small `A2:K16` scraps), `XJD(品質_状態・区分・ﾌﾗｸﾞ)`, `XJD(品質_ﾘｽﾄﾎﾞｯｸｽ)`, plus the 規格/測定/ﾛｯﾄ管理台帳 sheets |
| `_受注出荷` | `XJB(受注)`, `SJB(受注)` | `XJZ(共通)`, `XJA(基準)`, `XJC(工程)`, `XJD(品質)`, `XJE(生産）` |

A sheet-name pattern gets this wrong in **both** directions, and the two failures do not look alike:

- **Too narrow, and silent.** A hard-coded `XJC(*`/`SJC(*` matches nothing at all in the 基準情報 and
  共通 copies. The loop simply never runs and a header-only index is written with no error — which
  reads downstream as "every ID unregistered", or, once subset, as no findings at all.
- **Too broad, and noisy.** The obvious per-WG substitution over-matches: `XJD(*` for 品質管理 also
  picks up `XJD(品質_状態・区分・ﾌﾗｸﾞ)` and `XJD(品質_ﾘｽﾄﾎﾞｯｸｽ)`, neither of which is a dictionary.
  Likewise `XJC(*` against the 品質管理 or 受注出荷 copy hits that file's small non-dictionary
  `XJC(工程)` scrap sheet. These at least fail loudly, on the header probe.

So walk every visible sheet, keep the ones whose header pair resolves, **skip** the ones whose does
not, and **throw when a whole file yields zero dictionary sheets**. Probe the header band before
pulling the sheet's values: `翻訳リスト` alone is 7,042 rows that would be read and discarded.

This also disposes of the closed-pattern trap the name-matching approach carried (`"SJC(*)"` failing
against `SJC(工程)再開発追加分 `, trailing space included) — there is no pattern left to close.

Rows where the `連番` cell is empty are group separators (`ﾒﾆｭｰ`, `機能`, `ﾎﾞﾀﾝ`, `帳票`, `項目` …)
or unused placeholder rows — skip them on that test alone; do not try to detect section headings by
their text.

Launch Excel with the PID guard from `xlsx-excel-com-dump.md` ("Record every EXCEL.EXE PID that
already exists BEFORE launching our own instance") — omitted here only to keep the script readable.
`CsvQ`, `IsGray`, `LiveText` and `Test-IndexFresh` above are assumed to be in scope.

```powershell
$src   = Get-Item -LiteralPath $DictPath
$idxDir = Join-Path $env:USERPROFILE '.claude\skills\_cache\reference-index'
New-Item -ItemType Directory -Force -Path $idxDir | Out-Null
$idxPath = Join-Path $idxDir ("idx_screen-item-dictionary_" + $src.BaseName + ".csv")
if (Test-IndexFresh $idxPath $src.FullName) { return $idxPath }   # reuse; open no Excel at all

$excel = New-Object -ComObject Excel.Application     # + the PID guard, see above
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false
$wb = $null

$L = New-Object System.Collections.Generic.List[string]
$L.Add('# index: screen-item-dictionary')
$L.Add("# source: $($src.Name)")
$L.Add("# source-mtime-utc: $($src.LastWriteTimeUtc.ToString('o'))")
$L.Add("# source-length: $($src.Length)")
$L.Add('# columns: id,name')
$nRec = 0; $nRetired = 0; $nRenamed = 0; $nSheets = 0

# EVERYTHING that touches the open workbook goes inside try/finally. The builder throws on three
# layout surprises, and without this the throw skips Close/Quit entirely — cleanup then depends on
# the host process tearing down and releasing the COM references. That does happen when the script
# runs as its own short-lived process (measured: the orphan Excel exits a few seconds later, so
# don't judge a leak on a 1-2 second window), but it is not something to rely on inside a
# long-lived session. Verified: on the injected-fault path Quit() runs and no orphan survives.
try {
$wb = $excel.Workbooks.Open($src.FullName, $true, $true)

foreach ($sh in $wb.Worksheets) {
    if ($sh.Visible -ne -1) { continue }

    $used = $sh.UsedRange
    $rowsN = $used.Rows.Count; $colsN = $used.Columns.Count
    $r0 = $used.Row - 1; $c0 = $used.Column - 1
    $rows = [Math]::Min($rowsN, 5000); $cols = [Math]::Min($colsN, 60)

    # Probe ONLY the header band first. $used.Value2 on a non-dictionary sheet is pure waste -
    # 翻訳リスト is 7,042 rows - and we cannot know it is one until the header fails to resolve.
    $hp = [Math]::Min($rows, 20)
    $hv = $sh.Range($sh.Cells.Item(1 + $r0, 1 + $c0),
                    $sh.Cells.Item($hp + $r0, $cols + $c0)).Value2
    $hr = 0; $gc = 0; $nc = 0
    for ($r = 1; $r -le $hp -and -not ($gc -and $nc); $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $x = $hv[$r,$c]; if ($null -eq $x) { continue }
            $s = ([string]$x).Trim()
            if ($s -eq '画面項目ID') { $hr = $r; $gc = $c }
            elseif ($s -eq '画面項目名' -and $hr -eq $r) { $nc = $c }
        }
    }
    # No header pair = not a dictionary sheet (改訂履歴, 翻訳リスト, Sheet1, and the non-dictionary
    # scraps the 品質管理/受注出荷 copies carry under other WGs' names). Skip it, do not throw.
    if (-not ($hr -and $gc -and $nc)) { continue }
    $nSheets++

    $v = $used.Value2
    if ($rowsN -gt 5000) { Write-Warning "$($sh.Name): $rowsN rows, capped at 5000 - index is incomplete" }
    $gc1 = $gc + 1                                   # the 連番 sub-column
    $sub = $v[($hr+1), $gc1]
    if ($null -eq $sub -or ([string]$sub).Trim() -ne '連番') {
        throw "unexpected dictionary layout in $($sh.Name): expected '連番' sub-header, got '$sub'"
    }

    $uStrike = $used.Font.Strikethrough      # NOTE: not $ws/$wS - PowerShell variable names are
    $uColor  = $used.Font.Color              # case-INsensitive, so $wS would clobber a $ws sheet var
    $clean = ($uStrike -isnot [System.DBNull]) -and ($uStrike -eq $false) -and
             ($uColor  -isnot [System.DBNull]) -and (-not (IsGray $uColor))

    $L.Add("# sheet: $($sh.Name)")
    for ($r = $hr + 2; $r -le $rows; $r++) {
        $sfx = $v[$r,$gc1]; if ($null -eq $sfx) { continue }        # separator / unused row
        $pre = [string]$v[$r,$gc]
        if (($pre + ([string]$sfx)).Trim().Length -eq 0) { continue }

        if ($clean) {                                              # no formatting to resolve
            $L.Add((CsvQ (($pre + ([string]$sfx)).Trim())) + "," + (CsvQ $v[$r,$nc]))
            $nRec++
            continue
        }

        $ar = $r + $r0
        $a = LiveText $sh.Cells.Item($ar, ($gc  + $c0))
        $b = LiveText $sh.Cells.Item($ar, ($gc1 + $c0))
        if ($null -eq $b) { $nRetired++; continue }                # 連番 withdrawn = entry retired
        # The category cell struck while the sequence survives means the entry is mid-recategorisation.
        # Emitting the bare 連番 would put an id in the index that matches nothing in any design doc,
        # so treat it as retired rather than half-registered. (0 rows on 工程管理 today - a guard.)
        if ($null -eq $a -and ([string]$pre).Trim().Length -gt 0) { $nRetired++; continue }
        $id = (([string]$a) + ([string]$b)).Trim()
        if ($id.Length -eq 0) { $nRetired++; continue }

        $nmCell = $sh.Cells.Item($ar, ($nc + $c0))
        $nmLive = LiveText $nmCell
        if (($nmCell.Font.Strikethrough -is [System.DBNull]) -and $null -ne $nmLive) { $nRenamed++ }

        # If the array value and the absolute-cell value disagree on an unstruck row, the offset is
        # wrong - fail loudly instead of emitting a shifted index.
        if ($null -ne $a -and $null -ne $b -and
            ($pre + ([string]$sfx)).Trim() -ne $id -and
            ($sh.Cells.Item($ar,($gc+$c0)).Font.Strikethrough -eq $false) -and
            ($sh.Cells.Item($ar,($gc1+$c0)).Font.Strikethrough -eq $false)) {
            throw "UsedRange offset mismatch on $($sh.Name) row $ar"
        }

        $L.Add((CsvQ $id) + "," + (CsvQ $nmLive))
        $nRec++
    }
}
# Zero dictionary sheets means the wrong workbook, or a layout change. Fail loudly: writing the
# header-only index here is the one outcome nothing downstream can distinguish from "clean".
if ($nSheets -eq 0) {
    throw "no dictionary sheet found in $($src.Name) - is this a 画面項目辞書 workbook?"
}
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $excel.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
[IO.File]::WriteAllLines($idxPath, $L, [Text.Encoding]::UTF8)
"$idxPath : $nRec records, $nRetired retired, $nRenamed renamed-in-place"
```

**PowerShell indexing gotcha.** `$v[$r,$gc+1]` parses as a *three*-index lookup `$v[$r,$gc,1]` and
throws `You cannot index into a 2 dimensional array`. Always parenthesise: `$v[$r,($gc+1)]`, or
precompute `$gc1` as above.

## Subset to the program before handing the index to an agent

The index is the artefact you cache; it is not necessarily what an agent should read. For one
program's REV, filter it to the IDs that program actually cites — 106 of 3,288 records on
`SXJCB147`, ~1,100 tokens instead of ~28,200 (measured with the earlier, narrower ID filter that the
paragraph below replaces; the token intersection can only widen a subset). Select those records
**keeping multi-line records whole**:

**Scan only the sheet dumps, never the whole dump directory.** `$dump\*.txt` also matches
`_DELETED_DIGEST.txt`, whose whole purpose is to hold the content the live dump *removed* — on
`SXJCB147` that digest contributes 166 IDs, **101 of which the program does not cite in any live
cell**. Subsetting on it re-imports the struck-through references that the live dump exists to keep
out, inflates the subset from 106 records to 207, and invites exactly the "reasoning about deleted
text" finding that `xlsx-excel-com-dump.md` warns about. Filter the `_`-prefixed auxiliary files out.

**Do not filter with a regex for the ID's shape.** The obvious `\b(XJC|SJC)[0-9]{4}\b` is wrong
three ways, silently every time:

- **Too narrow.** An ID is `ＩＤ`+`連番` joined, so `MENUXJCP00`, `XJCGXJC101A` and `SJCRSJC006` (the
  renamed-in-place example above) are all real IDs. On 工程管理, **179 of 3,289 records (5.4%)** have
  a form that pattern cannot match.
- **`\b` does not fire against Japanese.** .NET counts kana and kanji as word characters, so
  `管理No：XJC8046` and `【SJC8613】` match while `画面項目SJC0600の値` yields nothing at all.
- **It bakes in one WG's prefixes**, so it needs editing per index — exactly the mistake the routing
  section exists to prevent.

Intersect the dump's ASCII word-runs with the IDs the index already holds instead. The index is the
authority on what an ID looks like, so there is nothing to guess: a run bounded by Japanese text is
still extracted as its own run, a non-dictionary token like the table ID `TXJCM003` simply is not in
the index and cannot match, and the same code subsets every WG's index unchanged.

```powershell
$sheetDumps = (Get-ChildItem -LiteralPath $dump -Filter '*.txt' |
               Where-Object { -not $_.Name.StartsWith('_') }).FullName
if (-not $sheetDumps) { throw "no sheet dumps under $dump" }
$tok = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($f in $sheetDumps) {
    $text = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
    foreach ($m in [regex]::Matches($text, '[0-9A-Za-z_]{4,}')) { [void]$tok.Add($m.Value) }
}
$idx  = [IO.File]::ReadAllLines($indexPath, [Text.Encoding]::UTF8)

$keep = New-Object System.Collections.Generic.List[string]
$i = 0; $total = 0
while ($i -lt $idx.Count) {
    $line = $idx[$i]
    if ($line.StartsWith('#')) { $keep.Add($line); $i++; continue }
    # absorb continuation lines until the record's quotes balance
    $rec = $line; $i++
    while ((([regex]::Matches($rec,'"')).Count % 2) -eq 1 -and $i -lt $idx.Count) {
        $rec += "`r`n" + $idx[$i]; $i++
    }
    # first field, quote-aware
    if ($rec.StartsWith('"')) {
        $e = 1
        while ($e -lt $rec.Length) {
            if ($rec[$e] -eq '"') {
                if ($e+1 -lt $rec.Length -and $rec[$e+1] -eq '"') { $e += 2; continue }   # escaped ""
                break                                                                     # closing quote
            }
            $e++
        }
        $id = $rec.Substring(1, $e-1).Replace('""','"')
    } else {
        $c = $rec.IndexOf(',')
        $id = if ($c -ge 0) { $rec.Substring(0,$c) } else { $rec }
    }
    if ($tok.Contains($id)) { $keep.Add($rec) }
    $total++
}
# Count the FULL index here, from the index itself. Do not carry the builder's $nRec over: the
# builder short-circuits and returns early whenever the cached index is still fresh, so on the
# common path that variable was never assigned and the line would come out empty — leaving the
# line come out blank, which reads as a truncated or failed subset.
$keep.Add("# full-index-record-count: $total")
[IO.File]::WriteAllLines($subsetPath, $keep, [Text.Encoding]::UTF8)
```

Hand the agent **both paths** — the subset to read whole, the full index to grep. A record count
cannot stand in for the full index: it says nothing about whether any particular ID is in there.

And be explicit with the agent about what a miss means, because the intuitive reading is backwards.
The subset is the full index filtered by the IDs found in this program's dump, so an ID the program
cites that is missing from the subset already **is** the unregistered-ID finding — not "this program
does not use it". The full index is the confirmation step, not a separate question: one grep
separates "genuinely absent from the dictionary" (the finding) from "absent from the subset because
the ID pattern above did not match its form" (a subsetting gap — report it, but it is not a design
defect). Telling the agent to suppress subset misses disables the only finding this check produces,
since by construction every cited-and-registered ID is already in the subset.

## What still needs the full workbook

Do not index a reference file whose check genuinely reads prose: `05.ｼｽﾃﾑ共通設計書.xlsx` sections
cited by ※-notes, `07.共通項目取得.xlsx` delegation blocks (a get-item block is a structure, not a
lookup row), and `09.区分名称_step2.xlsx` group bodies. Those keep using the `xlsx-dumps` cache in
`xlsx-excel-com-dump.md`, as do the three masters listed as not-yet-indexed at the top of this file.
Indexing is for ID→attribute lookups only.
