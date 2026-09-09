# Reference-master indexes: read a lookup table, not a whole workbook

`xlsx-excel-com-dump.md` covers the workbook **under review**. This file covers the reference masters
a REV skill consults as ground truth — the screen-item dictionary, table layouts, the DB registry.
Those are large, they change rarely, and every check that touches them needs only two or three
columns out of forty-odd. Dumping them whole is the single most expensive thing left in a REV.

Measured on `82.画面項目辞書_工程管理.xlsx`: the full-sheet dump is **4,214,998 characters
(~1,770,000 tokens)**. The index below is **56,300 characters (~28,200 tokens)** — a 75x reduction,
98.7% — and the subset a single program actually needs (the ~104 IDs it cites) is **~640 tokens**.
For the 19 table layouts of one program, the layout index is 39,022 chars (~15,100 tokens) against
114,381 chars (~48,000 tokens) of raw dump, a 66% reduction.

Read this file when a check needs 画面項目辞書, ﾃｰﾌﾞﾙﾚｲｱｳﾄ, DB一覧 or 帳票一覧 — that is
`design-doc-internal-consistency`, `xlsx-db-column-check` and `design-doc-io-table-check`. The other
three single-program skills never touch a reference master and should not read this file.

## Format: sectioned CSV

One record per line, hierarchy carried by `#` comment lines rather than repeated columns:

```csv
# index: screen-item-dictionary
# source: 82.画面項目辞書_工程管理.xlsx
# source-mtime-utc: 2026-09-08T07:25:24.9069190Z
# source-length: 4462592
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
38,867 (+56%), **sectioned CSV 20,894 (-16%)**. Flat CSV is the worst of the three — do not "simplify"
the section headers into columns.

The bigger reason is correctness. A Markdown table cannot represent a value containing `|` or a
newline, and its ` | value | ` convention silently strips leading and trailing whitespace. In the
工程管理 dictionary's ID and name columns alone there are **2 values containing a newline and 18 with
significant leading/trailing whitespace** (`'品目ﾏｽﾀ無し及び特殊作業ﾘｽﾄ '`, `'  /  /  '`, `'  月  日'`).
Those trim to something else, and a name-match check then reports a mismatch that does not exist —
the exact false-finding shape this project has been bitten by before. CSV with proper quoting is
lossless for all of them.

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

## Struck-through and grayed-out entries are excluded

Same rule as the workbook dump: a retired dictionary entry or a deleted layout column must not count
as registered, or the check reports "this ID exists" / "this column exists" about something the
design has already withdrawn. Apply the same cascade — ask the whole `UsedRange` once, and only walk
cells when it comes back mixed — and treat a `DBNull` (partially struck) cell as dead rather than
live. A half-struck dictionary entry is an edit in progress, not a usable ground truth.

```powershell
function IsGray([double]$argb) {
    $r = [int]$argb -band 0xFF; $g = ([int]$argb -shr 8) -band 0xFF; $b = ([int]$argb -shr 16) -band 0xFF
    return (($r -eq $g) -and ($g -eq $b) -and ($r -gt 80) -and ($r -lt 220))
}
function IsDead($cell) {
    $s = $cell.Font.Strikethrough
    if ($s -is [System.DBNull]) { return $true }   # mixed = mid-edit, exclude
    if ($s -eq $true) { return $true }
    $c = $cell.Font.Color
    if (-not ($c -is [System.DBNull])) { if (IsGray $c) { return $true } }
    return $false
}
```

Measured cost of this on real files: 2 excluded entries out of 3,287 in the 工程管理 dictionary, 1
column out of 1,045 across 19 table layouts. Cheap, and it removes a whole class of wrong answer.

## Freshness: rebuild when the source file changes

Every index carries `# source-mtime-utc` and `# source-length` from the source workbook. Before
using an index, stat the source and compare both. If either differs, or the index does not exist,
rebuild it; otherwise reuse it and open no Excel at all. This is the same contract as the
`xlsx-dumps` cache in `xlsx-excel-com-dump.md`, and it matters more here — these masters are edited
during a review cycle, and a stale dictionary produces "unregistered ID" findings for IDs somebody
added yesterday. Never reuse an index without the check.

Cache root: `<user home>\.claude\skills\_cache\reference-index\`. One `.csv` per source file, named
`idx_<kind>_<source stem>.csv`.

```powershell
function Test-IndexFresh($indexPath, $sourcePath) {
    if (-not (Test-Path $indexPath)) { return $false }
    $src = Get-Item -LiteralPath $sourcePath
    $head = Get-Content -LiteralPath $indexPath -TotalCount 8 -Encoding UTF8
    $m = ($head | Where-Object { $_ -like '# source-mtime-utc: *' }) -replace '^# source-mtime-utc: ',''
    $l = ($head | Where-Object { $_ -like '# source-length: *' })    -replace '^# source-length: ',''
    return ($m -eq $src.LastWriteTimeUtc.ToString('o')) -and ($l -eq [string]$src.Length)
}
```

## Builder: screen-item dictionary (`82.画面項目辞書_<WG>.xlsx`)

**Resolve columns by label, never by position — the sheets disagree with each other.** In
`82.画面項目辞書_工程管理.xlsx`, `XJC(工程)` has its 画面項目ID group at column 3 and `SJC(工程)再開発追加分 `
has it at column 4; everything downstream is shifted by one. A fixed-column reader silently produces
garbage on one of the two.

**The ID is two columns concatenated.** Row 6 carries the group headers (`画面項目ID`, `画面項目名`),
row 7 the sub-headers (`ＩＤ`, `連番`, `日本語`, `ベトナム語`). The 画面項目ID group spans the `ＩＤ`
column (a category: `MENU`, `XJC`, `SJC`) and the `連番` column (the rest: `8046`, `GXJC101A`,
`XJCP00`). What a design doc actually cites is the two joined: `XJC` + `8046` = `XJC8046`,
`MENU` + `XJCP00` = `MENUXJCP00`. Reading either column alone gives you IDs that appear nowhere in
any design doc — confirmed by building it wrong first, which produced a 3,349-row index full of
`MENU,工程管理` rows and no `XJC8046` at all.

Rows where the `連番` cell is empty are group separators (`ﾒﾆｭｰ`, `機能`, `ﾎﾞﾀﾝ`, `帳票`, `項目` …)
or unused placeholder rows — skip them on that test alone; do not try to detect section headings by
their text.

```powershell
$sh = ...                                    # a XJC(*/SJC(* sheet
$u = $sh.UsedRange; $v = $u.Value2
$rows = [Math]::Min($u.Rows.Count, 5000); $cols = [Math]::Min($u.Columns.Count, 60)

# header row + group columns, by label
$hr = 0; $gc = 0; $nc = 0
for ($r = 1; $r -le [Math]::Min($rows,20) -and -not ($gc -and $nc); $r++) {
    for ($c = 1; $c -le $cols; $c++) {
        $x = $v[$r,$c]; if ($null -eq $x) { continue }
        $s = ([string]$x).Trim()
        if ($s -eq '画面項目ID') { $hr = $r; $gc = $c }
        elseif ($s -eq '画面項目名' -and $hr -eq $r) { $nc = $c }
    }
}
$gc1 = $gc + 1                               # the 連番 sub-column
$sub = $v[($hr+1), $gc1]
if ($null -eq $sub -or ([string]$sub).Trim() -ne '連番') { throw "unexpected dictionary layout in $($sh.Name)" }

$wS = $u.Font.Strikethrough; $wC = $u.Font.Color
$clean = ($wS -isnot [System.DBNull]) -and ($wS -eq $false) -and
         ($wC -isnot [System.DBNull]) -and (-not (IsGray $wC))

for ($r = $hr + 2; $r -le $rows; $r++) {
    $sfx = $v[$r,$gc1]; if ($null -eq $sfx) { continue }      # separator / unused row
    $id = (([string]$v[$r,$gc]) + ([string]$sfx)).Trim()
    if ($id.Length -eq 0) { continue }
    if (-not $clean) {
        if ((IsDead $sh.Cells.Item($r,$gc)) -or (IsDead $sh.Cells.Item($r,$gc1)) -or
            (IsDead $sh.Cells.Item($r,$nc))) { continue }
    }
    $line = (CsvQ $id) + "," + (CsvQ $v[$r,$nc])
    # ... append to the sheet's section
}
```

**PowerShell indexing gotcha.** `$v[$r,$gc+1]` parses as a *three*-index lookup `$v[$r,$gc,1]` and
throws `You cannot index into a 2 dimensional array`. Always parenthesise: `$v[$r,($gc+1)]`, or
precompute `$gc1`. This failed silently enough on the first attempt to be worth stating.

## Builder: table layouts (`<TableID>_<name>.xlsx`, sheet `ﾃｰﾌﾞﾙﾚｲｱｳﾄ`)

Fixed template: table ID at `A6`, table name at `F6`, header at row 7, one column per row from row 8:
`[r,3]` = 項目名, `[r,12]` = 項目ID, `[r,19]` = type, `[r,23]` = length.

**Resolve the file by the ID inside it, never by filename.** A glob of `<TableID>_*.xlsx` also matches
suffixed *different* tables, because `_`-suffixed IDs are themselves real. Confirmed for real while
building this: resolving the 19 tables of `SXJCB147` by filename picked
`TXJCM003_B_ｵｰﾀﾞｰ投入出荷予定.xlsx` for `TXJCM003` (製造ｵｰﾀﾞｰ),
`TXJCM007_B_移動ﾛｯﾄ構成取消履歴.xlsx` for `TXJCM007` (移動ﾛｯﾄ構成) and
`TXJAM008_B_共通ｺｰﾄﾞﾏｽﾀ(選択肢).xlsx` for `TXJAM008` (共通ｺｰﾄﾞﾏｽﾀ) — three of nineteen pointing at the
wrong table's column list, which is a false "column does not exist" finding waiting to happen. Open
each candidate, read `A6`, and accept only an exact match. `TXJAM008_IN_共通ｺｰﾄﾞﾏｽﾀ受信.xlsx`
(`TXJAM008_IN`) is a fourth trap on the same table.

Candidate order stays PH3 > PH2 > top level for a WG folder, then
`01_Doc\07_データベース・ファイル設計書\` for 基準情報(JA)-prefix common tables — but the A6 check
overrides the order: a higher-priority file whose A6 disagrees is not the file.

```csv
# index: table-layout
# columns: colname,colid,type,len
# table: TXJCM137
# name: 処置指示ﾏｽﾀ
# file: TXJCM137_処置指示ﾏｽﾀ.xlsx
# source-mtime-utc: 2026-09-08T07:25:24.9069190Z
# source-length: 184320
会社ｺｰﾄﾞ,COMPANYCODE,NVARCHAR2,5
登録者,REGISTEREDPERSON,NVARCHAR2,20
登録日時,REGISTEREDDT,TIMESTAMP WITH TIME ZONE,
```

## Subset to the program before handing an index to an agent

The index is the artefact you cache; it is not necessarily what an agent should read. For one
program's REV, filter it to the IDs that program actually cites — 104 of 3,287 rows on `SXJCB147`,
~640 tokens instead of ~28,200. Grep the live dump for the ID pattern, then select those rows:

```powershell
$used = Select-String -Path "$dump\*.txt" -Pattern '\b(XJC|SJC)[0-9]{4}\b' -AllMatches |
        ForEach-Object { $_.Matches } | ForEach-Object { $_.Value } | Sort-Object -Unique
$idx  = [IO.File]::ReadAllLines($indexPath, [Text.Encoding]::UTF8)
$keep = $idx | Where-Object { $_.StartsWith('#') -or ($used -contains ($_ -split ',')[0]) }
```

Hand the agent the subset **plus** the count of rows in the full index, so it can tell "not in the
subset because this program does not use it" from "not in the dictionary at all". An ID the program
cites that is missing from the subset is the finding; an ID missing from the *index* is a different,
stronger finding.

## What still needs the full workbook

Do not index a reference file whose check genuinely reads prose: `05.ｼｽﾃﾑ共通設計書.xlsx` sections
cited by ※-notes, `07.共通項目取得.xlsx` delegation blocks (a get-item block is a structure, not a
lookup row), and `09.区分名称_step2.xlsx` group bodies. Those keep using the `xlsx-dumps` cache in
`xlsx-excel-com-dump.md`. Indexing is for ID→attribute lookups only.
