# Reading Unicorn design-doc Excel files (no Python/Node available)

This machine has no working Python or Node.js (`python3`/`python`/`py`/`node` all resolve to inert
Windows Store stub launchers). `unzip` is available in the Bash tool but raw xlsx XML/shared-strings
inspection is tedious. **Microsoft Excel is installed and reachable via COM automation from
PowerShell** — this is the reliable, fast way to read `.xlsx` contents in this environment. All
`*review*`/`*column-check*` skills in this skills folder should use this same technique.

## Folders to always exclude from project-wide searches

When any skill searches broadly across the whole `Unicorn` project root (Glob/Grep for a table ID,
a design-doc filename, a DB layout file, etc.), always exclude these path patterns up front —
each confirmed out of scope by the user directly. A file found only under one of these should never
be treated as a real hit (e.g. a table's layout "found" only in `90_Branches/` is not the same as it
being found in the real WG folder structure).

**Quick reference — exclude these from every recursive search's scope:**

- `jira_slack_notifier/**`
- `90_Branches/**`
- `01_Doc/04_共通設計/90_JAGUR各管理台帳/**`
- `*_*WG/開発DDL作成用*/**` (any WG, any date suffix)
- `**/11_単体テスト/**/サンプルデータ/**`
- `**/11_単体テスト/**/参考データ/**`

**Why** (background, skip on a routine run):

- `jira_slack_notifier/` — unrelated internal tooling repo, not a design-doc/DB-layout folder.
- `90_Branches/` — personal/WIP branch copies, not authoritative.
- `01_Doc/04_共通設計/90_JAGUR各管理台帳/` — superseded pre-migration ledger copies.
  `.claude/settings.json` denies `Read` here, but Glob/Grep and the Excel-COM dump script (opens by
  raw path, bypassing the `Read` deny) still see these files — exclude explicitly regardless.
- `<WG>WG\開発DDL作成用<date>\` — always excluded; a former exception for cross-WG table-layout
  lookup no longer applies.
- `.../11_単体テスト/.../サンプルデータ/` and `.../参考データ/` — test-data files whose names match
  real table IDs (e.g. `JAGUR.TXJAM008.csv`); confirmed never real evidence across multiple reviews.

## Don't parallelize a large table/sheet list across your own sub-agents

However large the list (e.g. `PSJCO308_焼成ｻﾔ詰め組み.xlsx`'s 29 更新条件表 sheets, or its 2092-row
画面設計書), work through it yourself, sequentially — do not fan out to your own sub-agents to split
it. Confirmed real failure on `PSJCO308`: both `design-doc-io-table-check` and `xlsx-db-column-check`
split their 29 tables across 4-7 sub-agents each; several got broken/empty prompts and returned
nothing, others were still running when the parent's own turn ended (forcing the orchestrating
session to step in and force a synthesis), and the extra request volume materially contributed to
the session hitting its rate limit. None of it was faster or more thorough than working the list
directly — it only added failure points and cost.

If a list is too large for one pass, work it in batches within your own turn (e.g. 5-10 tables at a
time, repeat) and report any remaining gap explicitly as a coverage limitation — don't silently drop
scope, and don't spin up sub-agents to cover the rest. Only the orchestrating session decides whether
to split a REV across multiple parallel agents.

## Orchestrating session: dump once, share the text, across parallel agents on the same workbook

When several sibling skills (`design-doc-internal-consistency`, `design-doc-io-table-check`,
`xlsx-db-column-check`, `naming-standard-compliance`, `design-doc-formatting-consistency`,
`design-doc-typo-check`) run as parallel background agents against the SAME target workbook, don't
let each one independently dump it from scratch — up to 6x duplicated Excel COM cycles on an
identical file. Confirmed real waste: a review of `XJC_ｼｽﾃﾑ共通設計書.xlsx`'s `ﾛｯﾄ停止ﾁｪｯｸ` sheet had
3 agents each re-dump the same 999-row sheet because their prompts didn't say to reuse an existing
dump.

Instead: before launching the batch, run the bulk `Value2` dump (per "The dump script" below) once
yourself against every sheet, and pass the resulting `.txt` file paths into each agent's prompt
("read `<path>` for sheet X — don't re-dump it"). Each agent still needs its own live Excel COM
access for anything formatting-based (strikethrough/gray-color, font-size/merge, per-character
DBNull) — a text dump can't carry those — so only the bulk `Value2` pass is worth sharing; don't try
to coordinate the formatting scans across agents too, that costs more than it saves.

## The dump script

Open the workbook invisibly, pull each worksheet's UsedRange as one `Value2` array (do **not**
loop `Cells.Item(r,c)` one cell at a time — it is slow enough to blow past the 2-minute Bash
timeout on any sheet with more than a few hundred cells), cap rows/cols since some sheets have a
bloated UsedRange from stray formatting far beyond the real content, and write one compact text
file per sheet with `[row,col]=value` tokens (skips empty cells, keeps coordinates for citing back
to the source file).

**Skip 詳細設計書* and *画面ｲﾒｰｼﾞ* sheets — don't dump their content at all.** Confirmed across every
REV skill's own instructions: 5 of the 6 single-program skills explicitly say "don't read/dump
詳細設計書 sheets" (it's near-empty boilerplate in practice), and `design-doc-formatting-consistency`
also excludes both 詳細設計書 and the 画面ｲﾒｰｼﾞ mockup sheet from its scan scope (the mockup is a
screenshot image with placeholder cells, confirmed content-free in every program checked so far).
Dumping their full cell content into the shared scratchpad every REV is pure waste — no sibling
skill ever reads those `.txt` files. **Do not extend this to every "ｲﾒｰｼﾞ"-named sheet** — a sheet
like `ﾒｰﾙｲﾒｰｼﾞ` (a batch program's mail-notification template) carries real narrative text that
`design-doc-internal-consistency`/`design-doc-typo-check` genuinely read and have found real
findings in (a stale placeholder wording, confirmed on `PXJCB134_流動停止(ﾊﾞｯﾁ).xlsx`) — only skip a
sheet whose name starts with `詳細設計` or contains `画面ｲﾒｰｼﾞ` specifically, not "ｲﾒｰｼﾞ" generally.

Record each skipped sheet's `UsedRange.Rows.Count`/`Columns.Count` (cheap — no `Value2` pull needed)
instead of a full dump, and report that alongside the rest. This one exception still needs it:
`naming-standard-compliance`'s "表紙's Ⅱ．設計書構成 list vs the sheets actually present" check needs
to tell whether a 詳細設計書 sheet is "populated with real content" or "an empty template" — the row
count alone is normally enough signal for that (a real 詳細設計書 runs dozens of rows; an untouched
template stub is a handful) without needing the full cell-value dump. If that skill's agent genuinely
needs to confirm actual content rather than just row count, it can open the one sheet itself via its
own light Excel COM check — that's still far cheaper than every REV dumping the full sheet by default.

**Format the `Value2` array via a compiled C# helper (`Add-Type`), not a plain PowerShell `for` loop.**
Benchmarked on `XJC_ｼｽﾃﾑ共通設計書.xlsx`'s `実績表項目設定` sheet (2652×116, ~307k cells): the
PowerShell loop took 7.8s to format vs 55ms for the equivalent compiled C# method — ~140x faster,
verified byte-identical output. The bottleneck is PowerShell interpreter overhead iterating cells in
memory, not the COM call (`Value2` itself pulls in 194ms) — so this costs nothing in correctness and
pays off most on the largest sheets. `ScreenUpdating`/`EnableEvents`/`Calculation` flags were also
tested and made no difference to `Workbooks.Open` time (~5.2-5.9s either way, fixed COM overhead) —
only the formatting loop is worth optimizing.

```powershell
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ExcelComWin32 {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
public static class XlsxDumpHelper {
    public static string FormatSheet(object vals, int rows, int cols) {
        var sb = new System.Text.StringBuilder();
        object[,] arr = vals as object[,];
        for (int r = 1; r <= rows; r++) {
            var parts = new System.Collections.Generic.List<string>();
            for (int c = 1; c <= cols; c++) {
                object v;
                if (arr == null) v = vals;
                else if (rows == 1) v = arr[1, c];
                else if (cols == 1) v = arr[r, 1];
                else v = arr[r, c];
                if (v != null) {
                    string s = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
                    if (s.Length > 0) parts.Add("[" + r + "," + c + "]=" + s);
                }
            }
            if (parts.Count > 0) {
                sb.Append(string.Join(" | ", parts));
                sb.Append("\r\n");
            }
        }
        return sb.ToString();
    }
}
"@

$out = "<scratchpad dir>"
# Record every EXCEL.EXE PID that already exists BEFORE launching our own instance. On this
# environment, `New-Object -ComObject Excel.Application` has been observed to sometimes attach to
# the user's own already-running interactive Excel process instead of spawning a genuinely new one
# (a Running-Object-Table quirk). If that happens and we later blindly call $excel.Quit(), it closes
# the user's ENTIRE real Excel session — including unrelated workbooks they had open — without
# saving. This bit it us once already: a REV skill's dump step silently attached to the user's
# session and Quit() force-closed a workbook they had open for unrelated work, mid-dump.
$preExistingExcelPids = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# Record the PID of THIS Excel instance only, via its (hidden) main window handle — never kill
# by process name later, since that would also hit Excel windows the user has open for other work.
[uint32]$excelComPid = 0
[ExcelComWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelComPid) | Out-Null
# CRITICAL: if the PID we got back was already running before we called New-Object, we attached to
# the user's live session rather than creating a new hidden one. Abort immediately — do NOT proceed
# to open/close/Quit anything on this $excel object, since every one of those calls would act on the
# user's real, visible Excel instance and their other open workbooks.
if ($preExistingExcelPids -contains $excelComPid) {
    throw "Excel COM automation attached to the user's existing Excel process (PID $excelComPid) instead of creating a new instance. Aborting without calling Open/Close/Quit on it — ask the user to close their other Excel windows first, or investigate why New-Object is reusing the running instance."
}
$wb = $excel.Workbooks.Open($path, $true, $true)   # ReadOnly, no update-links prompt
$skippedSheets = @()
foreach ($ws in $wb.Worksheets) {
    if ($ws.Name -like '詳細設計*' -or $ws.Name -like '*画面ｲﾒｰｼﾞ*') {
        $usedSkip = $ws.UsedRange
        $skippedSheets += "$($ws.Name): rows=$($usedSkip.Rows.Count) cols=$($usedSkip.Columns.Count) (skipped — out of scope for every REV skill; not dumped)"
        continue
    }
    $file = Join-Path $out ("<prefix>_" + ($ws.Name -replace '[\\/:*?"<>|]','_') + ".txt")
    $used = $ws.UsedRange
    $rows = [Math]::Min($used.Rows.Count, 3000)
    $cols = [Math]::Min($used.Columns.Count, 160)
    $vals = $used.Value2
    $text = [XlsxDumpHelper]::FormatSheet($vals, $rows, $cols)
    [System.IO.File]::WriteAllText($file, $text, [System.Text.Encoding]::UTF8)
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
$skippedSheets -join "`n"   # review this: a suspiciously large row count on a "skipped" 詳細設計書 sheet is a signal naming-standard-compliance may need to check its content directly
```

## Cross-session cache for reference/master files (not the target workbook)

The "dump once, share the text" section above only dedupes work *within* one REV run, for the one
target workbook under review. It does nothing for the other files a REV skill reads as ground
truth — `82.画面項目辞書_*.xlsx`, `04.ﾒｯｾｰｼﾞ管理_*.xlsx`, `09.区分名称_step2.xlsx`,
`05.ｼｽﾃﾑ共通設計書.xlsx`, `06-*.xlsx` (機能一覧/DB一覧/レスポンス一覧), `07.共通項目取得.xlsx`, the
チェックリスト file, and every テーブルレイアウト workbook. These almost never change between one
program's REV and the next, yet without caching they get re-opened via Excel COM from scratch every
single time, by every agent that needs them — including more than once within the same REV run, when
several sibling skills happen to need the same reference file.

**Use a persistent, cross-session cache keyed by the source file's last-write-time for any file in
this category.** Cache root: `<user home>\.claude\skills\_cache\xlsx-dumps\<md5 of the lowercased
absolute source path, plus an optional sheet-filter suffix — see below>\`, holding one
`<sheetName>.txt` per dumped sheet (same `[row,col]=value` format as above) plus a `meta.json`
recording the source path, `LastWriteTimeUtc`, file length, and which sheet filter (if any) was
used. Before dumping a reference file, compute its current `LastWriteTimeUtc`/length and compare
against `meta.json`: if both match, reuse the cached `.txt` files as-is (no Excel COM call at all —
this is the common case, since these files rarely change); if either differs, or `meta.json` doesn't
exist yet, redump via Excel COM (same safety-guarded approach as "The dump script" above — pre-existing-PID
check included) and overwrite the cache with the new dump plus updated `meta.json`. **Don't wipe the
cache directory with `Remove-Item -Recurse -Force` before redumping** — it isn't needed (the
directory is keyed by path + sheet-filter together, so the same cache dir always wants the exact
same set of sheet files every time it's redumped, and `Set-Content`/`WriteAllText` overwrite existing
files in place cleanly) and a recursive delete right next to a literal design-doc filename in the
same script has tripped this environment's own destructive-operation safety guard. Just
`New-Item -ItemType Directory -Force` (idempotent whether or not the folder already exists) and let
the per-sheet writes overwrite in place.

**Optional: restrict the dump to only the sheets a WG-scoped lookup can ever actually use, via
name patterns.** Two confirmed cases:

- **テーブルレイアウト files** — a table-layout workbook commonly has 4 sheets (checked
  `TXJCM501_ﾛｯﾄ停止.xlsx`): `改訂履歴`, `ﾃｰﾌﾞﾙﾚｲｱｳﾄ` (the one every skill actually reads),
  `JAG_ﾃｰﾌﾞﾙﾚｲｱｳﾄ` (an old JAGUR-era copy), `旧ﾃｰﾌﾞﾙﾚｲｱｳﾄ` (even older) — only 1 of 4 sheets is ever
  read. Use `$OnlySheetPatterns = @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ")`.
- **`04.ﾒｯｾｰｼﾞ管理_<WG>.xlsx` and `82.画面項目辞書_<WG>.xlsx`** — confirmed for real on the 工程管理
  copies: `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx` has 11 sheets (改訂履歴, 方針, XSA(共通), XSK(共通), XJZ(共通),
  XJA(基準), XJB(受注), **XJC(工程)**, **SJC(工程)**, XJD(品質), XJE(生産)) but the skills' own routing
  rules (see `design-doc-internal-consistency` check 3/4) mean only the reviewed program's own-WG
  JOBコード sheets (bold above) are ever read from THIS file — a common (共通) prefix routes to
  `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` instead, and a foreign-WG prefix routes to that other WG's own file, never
  to a same-named sheet embedded in this one. `82.画面項目辞書_工程管理.xlsx` shows the same shape (6
  sheets: SJC(工程)再開発追加分, Sheet1, **XJC(工程)**, 改訂履歴, 翻訳リスト(各Ver), 翻訳リスト — only
  the bold one plus the WG's own S-prefix sheet are read). Use
  `$OnlySheetPatterns = @("XJC(*", "SJC(*")` (or the equivalent for the WG you're checking) —
  `-like` wildcard patterns, not exact names, since a WG's own-prefix sheet can carry an extra
  suffix (confirmed: the real sheet name is `SJC(工程)再開発追加分 ` — extra text AND a trailing
  space — not the plain `SJC(工程)` you might expect) that an exact-name match would miss. **Leave
  the pattern open-ended with no closing `)`** — `-like` requires a literal `)` to match the actual
  end of the string, so a closed pattern like `"SJC(*)"` fails against `SJC(工程)再開発追加分 `
  (it doesn't end in `)`) even though it very much should match; confirmed this exact bug while
  testing against the real file. `"SJC(*"` (open-ended) matches both the plain and suffixed forms
  without accidentally matching an unrelated `XJC(...)` sheet.

Leave `$OnlySheetPatterns` `$null` (the default) for file types with no such known-safe restriction
(`09.区分名称_step2.xlsx`, `05.ｼｽﾃﾑ共通設計書.xlsx`, `06-*.xlsx`, `07.共通項目取得.xlsx`, the
checklist file, `04.ﾒｯｾｰｼﾞ管理_共通.xlsx`/`82.画面項目辞書_共通.xlsx` — these either have few sheets
already or every sheet genuinely gets read by some check). If none of the given patterns match any
sheet in a given workbook (a doc-shape exception), the script automatically falls back to dumping
every sheet and says so — it never silently drops content.

**Important: do not save this as a standalone `.ps1` file and invoke it with `-File` or dot-sourcing
— this environment's endpoint security (Cylance Script Control) blocks executing a `.ps1` file
outright, even via dot-sourcing.** Always paste the script inline as the PowerShell tool's command,
exactly like "The dump script" above.

```powershell
$SourcePath = "<absolute path to the reference file>"
$OnlySheetPatterns = $null   # e.g. @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ") or @("XJC(*","SJC(*"); $null to dump every sheet (note: no closing ")" on prefix patterns — see below)
$CacheRoot = Join-Path $env:USERPROFILE ".claude\skills\_cache\xlsx-dumps"

$resolved = Resolve-Path -LiteralPath $SourcePath
$SourcePath = $resolved.Path
$sourceInfo = Get-Item -LiteralPath $SourcePath
$currentMTime = $sourceInfo.LastWriteTimeUtc.ToString("o")
$currentLength = $sourceInfo.Length
$filterKey = if ($OnlySheetPatterns) { ($OnlySheetPatterns | Sort-Object) -join ";" } else { "" }

$md5 = [System.Security.Cryptography.MD5]::Create()
try { $hashBytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($SourcePath.ToLowerInvariant() + "|" + $filterKey)) }
finally { $md5.Dispose() }
$hash = -join ($hashBytes | ForEach-Object { $_.ToString("x2") })

$cacheDir = Join-Path $CacheRoot $hash
$metaPath = Join-Path $cacheDir "meta.json"

$cacheValid = $false
if (Test-Path $metaPath) {
    try {
        $meta = Get-Content -Raw -Encoding UTF8 $metaPath | ConvertFrom-Json
        if ($meta.sourcePath -eq $SourcePath -and $meta.lastWriteTimeUtc -eq $currentMTime -and $meta.length -eq $currentLength) {
            $cacheValid = $true
        }
    } catch { $cacheValid = $false }
}

if ($cacheValid) {
    Get-ChildItem -LiteralPath $cacheDir -Filter "*.txt" | ForEach-Object {
        Write-Output "$([System.IO.Path]::GetFileNameWithoutExtension($_.Name))`t$($_.FullName)"
    }
    Write-Output "CACHE_HIT: $cacheDir"
} else {
    New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null

    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ExcelComWin32Cache {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
public static class XlsxDumpHelperCache {
    public static string FormatSheet(object vals, int rows, int cols) {
        var sb = new System.Text.StringBuilder();
        object[,] arr = vals as object[,];
        for (int r = 1; r <= rows; r++) {
            var parts = new System.Collections.Generic.List<string>();
            for (int c = 1; c <= cols; c++) {
                object v;
                if (arr == null) v = vals;
                else if (rows == 1) v = arr[1, c];
                else if (cols == 1) v = arr[r, 1];
                else v = arr[r, c];
                if (v != null) {
                    string s = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
                    if (s.Length > 0) parts.Add("[" + r + "," + c + "]=" + s);
                }
            }
            if (parts.Count > 0) {
                sb.Append(string.Join(" | ", parts));
                sb.Append("\r\n");
            }
        }
        return sb.ToString();
    }
}
"@
    $preExistingExcelPids = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    [uint32]$excelComPid = 0
    [ExcelComWin32Cache]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelComPid) | Out-Null
    if ($preExistingExcelPids -contains $excelComPid) {
        throw "Excel COM automation attached to the user's existing Excel process (PID $excelComPid). Aborting without calling Open/Close/Quit on it."
    }

    $wb = $excel.Workbooks.Open($SourcePath, $true, $true)

    $targetSheets = @()
    if ($OnlySheetPatterns) {
        foreach ($ws in $wb.Worksheets) {
            foreach ($pat in $OnlySheetPatterns) {
                if ($ws.Name -like $pat) { $targetSheets += $ws; break }
            }
        }
        if ($targetSheets.Count -eq 0) {
            Write-Output "FALLBACK_NO_SHEET_MATCHED: none of the patterns ($($OnlySheetPatterns -join ', ')) matched any sheet in this workbook — dumping every sheet instead"
            $targetSheets = @($wb.Worksheets)
        }
    } else {
        $targetSheets = @($wb.Worksheets)
    }

    $dumpedSheetNames = @()
    foreach ($ws in $targetSheets) {
        $dumpedSheetNames += $ws.Name
        $safeName = ($ws.Name -replace '[\\/:*?"<>|]', '_')
        $file = Join-Path $cacheDir ("$safeName.txt")
        $used = $ws.UsedRange
        $rows = [Math]::Min($used.Rows.Count, 3000)
        $cols = [Math]::Min($used.Columns.Count, 160)
        $vals = $used.Value2
        $text = [XlsxDumpHelperCache]::FormatSheet($vals, $rows, $cols)
        [System.IO.File]::WriteAllText($file, $text, [System.Text.Encoding]::UTF8)
        Write-Output "$($ws.Name)`t$file"
    }
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    @{
        sourcePath = $SourcePath; lastWriteTimeUtc = $currentMTime; length = $currentLength
        onlySheetPatterns = $OnlySheetPatterns; dumpedSheets = $dumpedSheetNames
        dumpedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
    } | ConvertTo-Json | Set-Content -Path $metaPath -Encoding UTF8

    Write-Output "CACHE_MISS_REDUMPED: $cacheDir"
}
```

Read the resulting `<sheetName>.txt` files with the Read tool exactly as you would a fresh dump —
the cache is transparent to every downstream check in this document; only the redump step is
skipped when the source file hasn't changed. **This mechanism is specifically for reference/master
files, not the target design-doc workbook under review** — that one keeps using the plain "dump
once, share the text" flow above (it's reviewed once and the text is only needed for the duration of
that one REV, so a persistent cache buys nothing there and would only grow the cache directory with
one-off entries).

### Batch variant: checking many reference files in ONE Excel session

**Use this instead of the single-file version above whenever you already know you need to check more
than a couple of reference files up front** — most commonly, every テーブルレイアウト file for the
tables in a program's Ⅲ．入出力定義 list (`xlsx-db-column-check`, easily 10+ tables), or every
テーブルレイアウト file in a WG folder (`db-design-cross-consistency`, easily dozens). The single-file
script launches and quits a whole `Excel.Application` COM instance (Add-Type, `New-Object`, the
pre-existing-PID safety check, `Quit`/`ReleaseComObject`) per file — real, non-trivial overhead that
multiplies by file count if you paste it once per file. The batch variant checks every file's cache
validity FIRST, with no Excel interaction at all, and only launches a single shared `Excel.Application`
for the whole batch if at least one file actually needs a redump — so once the cache is warm (a
second REV of the same WG, or `xlsx-db-column-check` re-checking a table `db-design-cross-consistency`
already dumped), this typically launches Excel zero times.

```powershell
$Files = @(
    @{ SourcePath = "<absolute path to reference file 1>"; OnlySheetPatterns = @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ") }
    @{ SourcePath = "<absolute path to reference file 2>"; OnlySheetPatterns = @("XJC(*","SJC(*") }
    # ... one entry per file; OnlySheetPatterns = $null for a file type that needs every sheet
)
$CacheRoot = Join-Path $env:USERPROFILE ".claude\skills\_cache\xlsx-dumps"

function Get-XlsxCacheInfo($SourcePath, $OnlySheetPatterns, $CacheRoot) {
    $resolved = Resolve-Path -LiteralPath $SourcePath
    $SourcePath = $resolved.Path
    $sourceInfo = Get-Item -LiteralPath $SourcePath
    $currentMTime = $sourceInfo.LastWriteTimeUtc.ToString("o")
    $currentLength = $sourceInfo.Length
    $filterKey = if ($OnlySheetPatterns) { ($OnlySheetPatterns | Sort-Object) -join ";" } else { "" }
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try { $hashBytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($SourcePath.ToLowerInvariant() + "|" + $filterKey)) }
    finally { $md5.Dispose() }
    $hash = -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
    $cacheDir = Join-Path $CacheRoot $hash
    [PSCustomObject]@{
        SourcePath = $SourcePath; OnlySheetPatterns = $OnlySheetPatterns
        CurrentMTime = $currentMTime; CurrentLength = $currentLength
        CacheDir = $cacheDir; MetaPath = Join-Path $cacheDir "meta.json"
    }
}

# Pass 1: resolve cache status for every file — zero Excel interaction
$plan = @()
foreach ($f in $Files) {
    $info = Get-XlsxCacheInfo -SourcePath $f.SourcePath -OnlySheetPatterns $f.OnlySheetPatterns -CacheRoot $CacheRoot
    $valid = $false
    if (Test-Path $info.MetaPath) {
        try {
            $meta = Get-Content -Raw -Encoding UTF8 $info.MetaPath | ConvertFrom-Json
            if ($meta.sourcePath -eq $info.SourcePath -and $meta.lastWriteTimeUtc -eq $info.CurrentMTime -and $meta.length -eq $info.CurrentLength) { $valid = $true }
        } catch { $valid = $false }
    }
    $plan += [PSCustomObject]@{ Info = $info; CacheValid = $valid }
}

# Emit every cache hit immediately, and collect the real misses
$toRedump = @()
foreach ($p in $plan) {
    if ($p.CacheValid) {
        Get-ChildItem -LiteralPath $p.Info.CacheDir -Filter "*.txt" | ForEach-Object {
            Write-Output "$([System.IO.Path]::GetFileNameWithoutExtension($_.Name))`t$($_.FullName)"
        }
        Write-Output "CACHE_HIT: $($p.Info.SourcePath)"
    } else {
        $toRedump += $p
    }
}

# Pass 2: only if at least one file is a genuine miss, launch ONE Excel instance for the whole batch
if ($toRedump.Count -gt 0) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ExcelComWin32Batch {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
public static class XlsxDumpHelperBatch {
    public static string FormatSheet(object vals, int rows, int cols) {
        var sb = new System.Text.StringBuilder();
        object[,] arr = vals as object[,];
        for (int r = 1; r <= rows; r++) {
            var parts = new System.Collections.Generic.List<string>();
            for (int c = 1; c <= cols; c++) {
                object v;
                if (arr == null) v = vals;
                else if (rows == 1) v = arr[1, c];
                else if (cols == 1) v = arr[r, 1];
                else v = arr[r, c];
                if (v != null) {
                    string s = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
                    if (s.Length > 0) parts.Add("[" + r + "," + c + "]=" + s);
                }
            }
            if (parts.Count > 0) {
                sb.Append(string.Join(" | ", parts));
                sb.Append("\r\n");
            }
        }
        return sb.ToString();
    }
}
"@
    $preExistingExcelPids = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    [uint32]$excelComPid = 0
    [ExcelComWin32Batch]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelComPid) | Out-Null
    if ($preExistingExcelPids -contains $excelComPid) {
        throw "Excel COM automation attached to the user's existing Excel process (PID $excelComPid). Aborting without calling Open/Close/Quit on it."
    }

    foreach ($p in $toRedump) {
        $info = $p.Info
        New-Item -ItemType Directory -Force -Path $info.CacheDir | Out-Null

        $wb = $excel.Workbooks.Open($info.SourcePath, $true, $true)
        $targetSheets = @()
        if ($info.OnlySheetPatterns) {
            foreach ($ws in $wb.Worksheets) {
                foreach ($pat in $info.OnlySheetPatterns) {
                    if ($ws.Name -like $pat) { $targetSheets += $ws; break }
                }
            }
            if ($targetSheets.Count -eq 0) {
                Write-Output "FALLBACK_NO_SHEET_MATCHED: none of the patterns ($($info.OnlySheetPatterns -join ', ')) matched any sheet in $($info.SourcePath) — dumping every sheet instead"
                $targetSheets = @($wb.Worksheets)
            }
        } else { $targetSheets = @($wb.Worksheets) }

        $dumpedSheetNames = @()
        foreach ($ws in $targetSheets) {
            $dumpedSheetNames += $ws.Name
            $safeName = ($ws.Name -replace '[\\/:*?"<>|]', '_')
            $file = Join-Path $info.CacheDir ("$safeName.txt")
            $used = $ws.UsedRange
            $rows = [Math]::Min($used.Rows.Count, 3000)
            $cols = [Math]::Min($used.Columns.Count, 160)
            $vals = $used.Value2
            $text = [XlsxDumpHelperBatch]::FormatSheet($vals, $rows, $cols)
            [System.IO.File]::WriteAllText($file, $text, [System.Text.Encoding]::UTF8)
            Write-Output "$($ws.Name)`t$file"
        }
        $wb.Close($false)

        @{
            sourcePath = $info.SourcePath; lastWriteTimeUtc = $info.CurrentMTime; length = $info.CurrentLength
            onlySheetPatterns = $info.OnlySheetPatterns; dumpedSheets = $dumpedSheetNames
            dumpedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        } | ConvertTo-Json | Set-Content -Path $info.MetaPath -Encoding UTF8

        Write-Output "CACHE_MISS_REDUMPED: $($info.SourcePath)"
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
```

Same freshness semantics as the single-file version (mtime + length + sheet-filter all folded into
the cache key), same Cylance constraint (paste inline, never save as a standalone `.ps1`), and the
same read-the-resulting-`.txt`-files-with-Read-tool usage afterward. Build the `$Files` list from
whatever table IDs/paths you already resolved (e.g. step 3's phase-precedence resolution in
`xlsx-db-column-check`) before running this once, rather than looping the single-file script
per table.

## Excluding struck-through / grayed-out rows from review

Design docs are often copied from an older workbook and edited in place. This project's writing-rule
checklist (`01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx`, sheet
"ﾁｪｯｸﾘｽﾄ", item 1-5) says authors should delete struck-through rows entirely and turn red
"to-be-fixed" text back to black when reusing a workbook this way — but that cleanup is often
skipped, leaving struck-through/grayed rows still physically present. **Content marked this way is
inactive/deleted-in-spirit and must never be treated as something the program actually
references** — don't count it toward a "referenced columns" list, don't flag it as a
naming/consistency violation, don't cite it as evidence of what the current design does. Skip it
silently (a one-line "N struck-through/grayed rows skipped" note in the final report is enough — don't
enumerate them as findings).

**Mandatory on every section you cite as evidence — never skip it for a large sheet.** A real review
(PXJCO193) reported six tables as "used but not declared in the I/O table," entirely because the
formatting scan was skipped on a 1000+ row sheet whose 参照ｴﾝﾃｨﾃｨ blocks were red-and-struck-through
100+ rows deep. Sheet size is exactly why the scan matters more, not a reason to shortcut it.

**This check is just as mandatory in the opposite direction — before reporting a section as a
"leftover that should have been deleted but wasn't cleaned up" (e.g. a stale JOIN/結合条件 block left
over after its referenced alias was swapped to a delegated sub-block that shouldn't need one), check
whether it is already struck through.** It's easy to run the formatting scan diligently everywhere you
build a "used" evidence list, and still skip it on a block you're about to flag as a *missing* deletion
— but a bulk `Value2` dump can't distinguish "still live, should be removed" from "already marked
dead, nothing to do" any more than it can distinguish live from dead evidence. Confirmed for real on
`XJC_ｼｽﾃﾑ共通設計書.xlsx`, sheet `ﾛｯﾄ停止ﾁｪｯｸ`, rows 829-832 (a "結合条件(A INNER JOIN B)" block left
over after its alias B was reassigned from a real table to a delegated get-item sub-block): a review
reported this as a live, not-yet-cleaned-up defect, but every cell in the block ([829,4] through
[832,28]) was actually `Font.Strikethrough=True` — it had already been correctly deleted, and the
"defect" was a pure false positive from not running the formatting check on it before writing up the
finding.

The bulk `Value2` dump above cannot see formatting, so this needs a second, *targeted* pass — keep it
targeted to avoid the same per-cell-loop slowness the bulk dump avoids. Once you've identified the
column that carries the item/label names for the section you care about, loop `Font.Strikethrough`
and `Font.Color` over just *that one column* for the row range in play (a few hundred COM calls, not
tens of thousands). This isn't limited to obvious "label" columns like 項目名 (更新条件表/
ﾃｰﾌﾞﾙﾚｲｱｳﾄ) or 画面項目名 (画面設計書) — it applies just as much to the table-name column inside a
"参照ｴﾝﾃｨﾃｨ" block (画面設計書 Ⅲ．画面表示仕様), since a struck-through query block's table
references are exactly what gets miscounted as "actively used" if formatting isn't checked first.
Any column you're about to cite as evidence needs this same scan before you trust it.

**Don't re-fetch `Value2` per row to test for emptiness — you already have it from the bulk dump.**
An earlier version of this script looped `$firstRow..$lastRow` and called `$cell.Value2` on every
row just to decide whether to skip it, which is exactly the per-cell-COM-call cost the bulk dump
exists to avoid, paid a second time on the same range. The bulk dump (or its `.txt` file, already
written before this scan ever runs) already tells you precisely which rows in this column are
non-empty — reuse that list and only spend COM calls on `Font.Strikethrough`/`Font.Color` for rows
you already know have content:

```powershell
$ws = $wb.Worksheets.Item("<sheet name>")
$col = 4          # the column holding the label text, e.g. 項目名

# Derive the known-non-empty row list for this column from the sheet's already-dumped .txt file
# (or from the bulk $vals array directly if still in scope from the same dump pass) — do NOT
# rediscover it by scanning Value2 via COM again.
$knownRows = Select-String -Path $dumpTxtPath -Pattern "\[(\d+),$col\]=" |
    ForEach-Object { [int]$_.Matches[0].Groups[1].Value } | Sort-Object -Unique

foreach ($r in $knownRows) {
    $cell = $ws.Cells.Item($r, $col)
    $strike = $cell.Font.Strikethrough
    $argb = $cell.Font.Color            # OLE color: 0x00BBGGRR
    $rr = $argb -band 0xFF
    $gg = ($argb -shr 8) -band 0xFF
    $bb = ($argb -shr 16) -band 0xFF
    $isGray = ($rr -eq $gg) -and ($gg -eq $bb) -and ($rr -gt 80) -and ($rr -lt 220)
    if ($strike -or $isGray) {
        Write-Output "[$r,$col] strike=$strike gray=$isGray"
    }
}
```

On a sparse column (say 20% filled over a 1000-row range), this cuts the COM-call count for this
scan by roughly half or more compared to the old per-row `Value2` re-check — the savings scale with
how sparse the column is, so they matter most on exactly the large sheets where this scan already
takes the longest.

Run this once per sheet you're extracting referenced items from, cross off any row number it
reports from the row set you built from the bulk dump, then proceed with the remaining rows as
normal.

**Trust this signal even when it covers an entire column, including the header cell and rows you
independently know are live** — per explicit user direction, red+strikethrough always means
"exclude from Rev scope" in this project, with no exception for how broadly it's applied. Do not
reason your way out of it ("this can't be meaningful, it's on every row including the header, must
be leftover noise") — that reasoning has been tried and explicitly overruled. If an entire
更新条件表/画面設計書 sheet's item column comes back red+strikethrough top to bottom, including its
own header, the correct read is that the **whole sheet/section is marked superseded/deprecated**
(commonly: an old JAGUR-era 更新条件表(<TableID>) sheet that a newer 更新条件表(<TableID>WF) sheet
has replaced) — exclude the whole thing from the review and say so, rather than treating the columns
inside it as individually-live references. Only a sheet/section with genuinely mixed formatting
(some rows marked, neighboring rows in the same column plainly not) should be filtered row-by-row;
a uniformly-marked sheet is excluded wholesale.

**A cell's `Font.Strikethrough` can itself come back as `DBNull` (type `System.DBNull`), exactly
like the already-documented `Font.Size` mixed-run case — this means the cell has *some* characters
struck through and others not, not that the check failed.** This happens whenever a stale reference
is edited in place instead of fully replaced — treating a DBNull result as "not struck" silently
keeps the stale part as if it were still live; treating it as "struck" (per the whole-column rule
below) would wrongly discard the live part too. Confirmed for real: `PSJCO205_返品処置指示発行.xlsx`,
`更新条件表(TXJCM006)`, rows 74/84/95, col 16 (取得元) — raw `③④` (a renumbered circled-entity
reference), whole-cell check DBNull, per-character inspection shows only `④` live (the sheet's own
row-9 legend documents the renumbering that stranded `③`). Same shape in a plain column-name cell:
`SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx`, `帳票設計書(RSJC035)`, `F202`/`Q202` — raw `注意事項備考`/
`A.注意事項備考`, only `注意事項` live; skipping the per-character check here produced a false
"referenced column 注意事項備考 doesn't exist in TXJCM137" finding (TXJCM137 has both `備考` and
`注意事項` as separate columns — the doc was correctly referencing the latter with a stale,
not-fully-deleted `備考` suffix glued on). Whenever a whole-cell strikethrough check returns DBNull,
drop to per-character inspection before trusting the cell's text as evidence:

```powershell
$cell = $ws.Cells.Item($r, $c)
$v = "$($cell.Value2)"
$whole = $cell.Font.Strikethrough
if ($whole -is [System.DBNull]) {
    $clean = ""
    for ($i = 1; $i -le $v.Length; $i++) {
        $ch = $cell.Characters($i, 1)
        if (-not $ch.Font.Strikethrough) { $clean += $ch.Text }
    }
    Write-Output "[$r,$c] mixed strikethrough: raw='$v' live-only='$clean'"
}
```

Use only the reconstructed `live-only` text as the cell's real content (e.g. treat `④` as the sole
source reference, not `③④`) — don't fall back to the raw `Value2` string once a cell has flagged as
mixed.

**Test the type explicitly (`-is [System.DBNull]`) — never eyeball a printed/interpolated value.**
`System.DBNull` interpolates into a PowerShell string as an *empty string*, indistinguishable at a
glance from a normal blank/false result in ad-hoc output like `Write-Output "strike=$whole"` (prints
`strike=` either way) — a review can run the whole-cell check, see nothing after `strike=`, read that
as "not struck, live," and move on without ever hitting the DBNull branch. Confirmed for real:
`PXJCO124_ﾛｯﾄ振向け.xlsx`, `ﾁｪｯｸ処理設計書(GXJC124A)`, `[280,17]` = `COUNT(A.層数層No)` — reported as
referencing a nonexistent column and "fixed" to `COUNT(A.層No)`, but `層No` was already the live text
(`層数` was the struck leftover); the doc was already correct.

**This same per-character check applies just as much to a single-cell "typo"/mismatch finding as to
reference-list evidence** — it's easy to apply diligently to referenced-column lists and still skip
it on a cell that just looks like a misspelling at a glance. Confirmed for real:
`PSJCO205_返品処置指示発行.xlsx`, `画面設計書(GSJC205A)`, `[333,28]` — raw `"GSXJC205A"` looks like a typo of
the real screen ID `GSJC205A` (extra `X`), reported as a confirmed defect, but per-character
inspection showed only the `X` struck through — live text is `"GSJC205A"`, an exact match, no defect.
Run the per-character check on any cell before including it as a mismatch/typo finding, same as
before building a referenced-items list.

**Watch for parallel `処理区分="..."の場合` (or similarly-named) branch variants inside one sheet** —
a common shape in this project's 画面設計書 is two or more parallel sub-blocks handling different
values of the same discriminator (e.g. `品目コード` vs `管理No`, one per branch), often followed by
several more sub-blocks that logically belong to just one of those branches. When one branch is
deprecated, the strikethrough frequently covers not just that one block but every subsequent block
that depended on it too, spanning hundreds of rows and several numbered sub-sections in a row. Seeing
one such branch header struck through is a strong signal to check the formatting of everything until
the next clearly-unstruck section header, rather than assuming the deprecation is scoped to just the
one block you first noticed it on.

## Font-size and cell-merge irregularity detection — moved

That technique is specific to `design-doc-formatting-consistency` and now lives in its own file,
`_shared/xlsx-formatting-scan.md`, so the other 5 REV skills don't pay the token cost of a section
they never use. Only `design-doc-formatting-consistency` needs to read that file.

## Program structure variants to expect

Every REV skill's procedure is written assuming the "default" shape — one プログラムID, one 画面ID,
one 画面設計書 sheet. Two real variants recur often enough in this project that they're worth
checking for up front, before assuming a section is simply missing or that a skill's default
evidence-source doesn't apply:

**A program can have NO 画面設計書 sheet at all** — this happens for a pure batch program (プログラムID
ends in `B`, e.g. `PXJCB102`, `PXJCB134`) or a report-only サブプロ (`S`-prefixed, e.g. `SXJCB147`).
Confirm this from 表紙's "Ⅱ．設計書構成" list (画面設計書 row marked `-`/`-`) before concluding
anything about screen-item checks doesn't apply. When there's no 画面設計書, the column-level
get-item evidence (参照ｴﾝﾃｨﾃｨ blocks, 取得項目/検索条件/結合条件) that a screen program would carry
in 画面設計書's "Ⅲ．画面表示仕様" instead lives in one of two places, and both need checking:
- **機能定義書's own "Ⅳ．機能処理概要"** — a pure batch program (no 帳票設計書 either, e.g.
  `PXJCB102`, `PXJCB134`) writes its 参照ｴﾝﾃｨﾃｨ blocks directly inline in this narrative section,
  numbered as sub-steps (e.g. "(2).共通ｺｰﾄﾞﾏｽﾀより完了予定日に加算する月数を取得する。" followed by
  its own 参照ｴﾝﾃｨﾃｨ/取得項目/検索条件 block).
- **Each 帳票設計書(<ReportID>) sheet's own get-item blocks** — a report-generating サブプロ
  (`SXJCB147` is the confirmed example) delegates almost all of its actual table access to its
  帳票設計書 sheets' own "Ⅲ．編集仕様" sections, each with the same 参照ｴﾝﾃｨﾃｨ/取得項目/検索条件/
  結合条件/ｿｰﾄ順 shape as a screen's get-item block — 機能定義書 itself may carry only one or two
  such blocks (or none) and just says "帳票設計書通りとする。" for the rest. Don't stop at "this
  program's 機能定義書 barely references any tables" — check every 帳票設計書 sheet too before
  concluding a declared table is unused, or that a table used in a report isn't declared.

**A single プログラムID can pair with MORE than one 画面ID** — confirmed for real on
`PSJCO304_着手ﾒｯｾｰｼﾞﾒﾝﾃﾅﾝｽ.xlsx`: one プログラムID (`PSJCO304`) pairs with two screens, `GSJC304A`
and `GSJC304B`, each with its own `画面設計書(GSJC304A)`/`画面設計書(GSJC304B)` sheet and its own
`ﾁｪｯｸ処理設計書(GSJC304A)`/`ﾁｪｯｸ処理設計書(GSJC304B)` sheet — while sharing a single set of
更新条件表 sheets and a single Ⅲ．入出力定義 table in one 機能定義書 sheet. Confirm the actual sheet
list before assuming "one 画面設計書" — if there's more than one, run every per-screen check (event↔
processing-overview, screen-item ID dictionary, message ID, three-way match, item-order consistency,
DB-column extraction) **separately for each screen**, and when verifying I/O-table usage, search
*all* screens' get-item blocks (not just the first one you find) before calling a table "unused."
Also expect a `ﾌｧｲﾙ出力仕様書(<FileID>)` sheet to show up alongside multi-screen programs like this —
treat its own get-item blocks (if any) the same way as a 帳票設計書's, per the batch-program note
above.

## 区分名称/区分コード lookups — always use the STEP2 file directly, and match the group name exactly

A design doc's ※-note frequently delegates a code→label lookup to `ｼｽﾃﾑ共通設計書.区分名称`,
naming a specific group (e.g. "区分名称 ﾛｯﾄ停止区分(ﾛｯﾄ停止指示登録) を参照"). **Always go straight
to `01_Doc/04_共通設計/09.区分名称_step2.xlsx`, sheet `区分名称_STEP2～`, as ground truth** — per
explicit user direction, this is a standing rule, not a per-run judgment call. Do not open the
unsuffixed `01_Doc/04_共通設計/09.区分名称.xlsx` at all (it's a superseded older version), and don't
probe the folder for some other/higher `_stepN` variant each time — STEP2 is the one to use, always.

**Independently of file version, a single 区分名称 file can carry more than one group with very
similar names — an old group and its renamed/expanded replacement coexisting side by side** — so a
keyword search alone (e.g. searching for "ﾛｯﾄ停止区分") can silently match the wrong one. Confirmed
for real: `09.区分名称_step2.xlsx`'s "区分名称_STEP2～" sheet contains BOTH a plain **`ﾛｯﾄ停止区分`**
group (row 846, an old code scheme: `4`=ﾃｰﾌﾟﾛｯﾄNo, `5`=波及範囲検索, `6`=製造ﾛｯﾄNo, no `C`/`D`) AND a
separately-named **`ﾛｯﾄ停止区分(ﾛｯﾄ停止指示登録)`** group (row 2249, the current scheme actually used
by `XJC_ｼｽﾃﾑ共通設計書.xlsx`'s `ﾛｯﾄ停止ﾁｪｯｸ` sheet: `4`=品目ｺｰﾄﾞ+ﾃｰﾌﾟﾛｯﾄNo, `5`=製造ﾛｯﾄNo,
`6`=品目ｺｰﾄﾞ+製造ﾛｯﾄNo(部分一致), through `C`/`D`). A review that matched on the shorter/plainer name
without checking the design doc's own ※-note for the EXACT full group name in parentheses reported a
"stale master, codes don't match" defect that was entirely a wrong-lookup false positive — the
correctly-named group matched the live design doc's codes perfectly. Always copy the group name
verbatim from the design doc's own delegation note (including any parenthesized qualifier) and match
it exactly against the 区分名称 sheet's group-header cells, rather than fuzzy/keyword-matching the
first similarly-named group found.

## Operational notes

- **Always check the dump's own row/col cap against the sheet's real size before trusting a "not
  found" result.** The template script's default cap is 3000 rows / 160 cols (raised from an
  earlier 500/100 default specifically because that was too low for this project's real sheets and
  caused a recurring wasted round-trip: dump at 500 → a section silently missing past row 500 →
  redump the same sheet at a higher cap → re-read. Confirmed sheet sizes that would have tripped the
  old cap: `PSJCO308`'s 画面設計書 (2092 rows), `XJC_ｼｽﾃﾑ共通設計書.xlsx`'s `実績表項目設定` (2652
  rows) and `ﾛｯﾄ停止ﾁｪｯｸ` (999 rows, 126 cols), `09.区分名称_step2.xlsx`'s `区分名称_STEP2～` (2371
  rows, 156 cols) — the new default covers all of these in one pass. Still, don't treat 3000/160 as
  a guarantee: re-check `$used.Rows.Count`/`$used.Columns.Count` from the sheet (printed when you
  dump it) against the cap, and if a sheet is bigger than even this default, redump just that sheet
  with an explicitly higher cap before concluding a section or ID is actually missing.
- Reuse a single `$excel` instance across many files (loop the whole batch inside one PowerShell
  call) instead of relaunching Excel per file — much faster, and this project has hundreds of
  design-doc workbooks per WG.
- `$excel.Visible = $false` and `$excel.DisplayAlerts = $false` before opening anything; always
  `$wb.Close($false)` per file so it doesn't prompt to save.
- If a run times out or errors, clean up the specific hidden instance this script started before
  retrying — a stale instance keeps file handles locked and blocks subsequent opens:
  `Stop-Process -Id $excelComPid -Force -ErrorAction SilentlyContinue`
  **Never run `Get-Process EXCEL | Stop-Process -Force`** (or any other kill-by-process-name
  form) — `EXCEL.EXE` is shared with any Excel window the user has open for unrelated work, and
  killing by name force-closes those too, discarding unsaved edits. Always target the exact PID
  captured via `$excelComPid` (or the equivalent for each instance, if a run opens more than one).
- **Always run the pre-existing-PID check in the dump script above before touching `$wb`/`$excel`
  at all.** This project has already hit the case where `New-Object -ComObject Excel.Application`
  attached to the user's own running Excel instead of creating a new hidden one — the PID-targeted
  kill guard above only protects the error-cleanup path, it does nothing for this failure mode,
  since the normal end-of-script `$wb.Close($false)` / `$excel.Quit()` calls are exactly what force-
  closed the user's unrelated open workbook the one time this happened. If the check throws, stop
  and tell the user rather than working around it.
- Then read the dumped `.txt` files with the Read tool (not `cat`/Bash) — the `[row,col]=value`
  format is compact enough to scan quickly and lets you cite exact cells back to the user.
- Table/column-layout workbooks in `07_データベース・ファイル設計書(仮)`-style folders follow a
  fixed template: a **"ﾃｰﾌﾞﾙﾚｲｱｳﾄ"** sheet with the table ID/name at row 6
  (`[6,1]=<TableID>` / `[6,6]=<Table name>`), header at row 7
  (`No. | 項目名 | 項目ID | 属性 | 桁数 | DB桁 | I01... | notnull | 備考`), and one data row per
  column from row 8 onward — `[r,3]` is the Japanese column name (項目名), `[r,12]` the physical
  column id, `[r,19]`/`[r,23]` the DB type/length.
- Program/screen design workbooks (機能定義書系) follow a fixed sheet set: 表紙 → 機能定義書 →
  画面設計書 → 詳細設計書 / ﾁｪｯｸ処理設計書 → 更新条件表(<TableID>) — section headers use Roman
  numerals (Ⅰ．Ⅱ．Ⅲ...) which are stable anchors to grep/read for across programs.
