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

Instead: before launching the batch, run the live dump (per "The dump script" below) once yourself
against every sheet, and pass the resulting `.txt` file paths into each agent's prompt ("read
`<path>` for sheet X — don't re-dump it").

**In the same pre-launch step, build the reference-master index if any selected check needs it —
see `_shared/reference-index.md`.** Today that means `design-doc-internal-consistency`, which reads
the 画面項目辞書 index. That doc tells the *agent* not to build the index and to stop if it is
missing, so if you skip this step the check simply does not run. Build one index per dictionary file
the program's IDs route to (that doc's routing table; a program using shared `XJZ`/`SJZ` items needs
the `_共通` file too), subset each to the program's own IDs, and pass both paths per file in the
prompt (the subset to read whole, the full index to grep) — same discipline as the dump.

**The dump script already applies strikethrough and gray-out, so a shared dump now carries
everything all but one skill needs — no agent should run its own strikethrough/gray scan.** Struck
cells are absent from the `.txt` entirely and partially-struck cells carry only their live text, so
"is this row still live?" is answered by whether it appears at all. Removed content is preserved
separately in `_DELETED_DIGEST.txt` for the one check that needs it (see "Excluding struck-through /
grayed-out rows from review" below). Measured on `SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx` (7 sheets): the
old shape cost ~155k tokens per agent (raw dump ~105k + strikethrough scan ~49k); the live dump
alone is ~89k (**-43%**), and live + digest is ~117k (-25%). Multiply that by 5-6 parallel agents.

The one remaining exception is `design-doc-formatting-consistency`, which needs `Font.Size` and
`MergeCells` — those still require live COM and are its own concern. Don't try to coordinate that
scan across agents; it costs more than it saves, and only one skill reads it.

## The dump script

Open the workbook invisibly, pull each worksheet's UsedRange as one `Value2` array (do **not**
loop `Cells.Item(r,c)` one cell at a time — it is slow enough to blow past the 2-minute Bash
timeout on any sheet with more than a few hundred cells), cap rows/cols since some sheets have a
bloated UsedRange from stray formatting far beyond the real content, and write one compact text
file per sheet with `[row,col]=value` tokens (skips empty cells, keeps coordinates for citing back
to the source file).

**The dump is a *live* dump: strikethrough and gray-out are resolved while the workbook is open, not
left for each agent to redo.** A fully-struck (or gray) cell is omitted from the `.txt`; a cell with
mixed struck/unstruck characters is written with only its live text; everything removed is written
to a separate `_DELETED_DIGEST.txt`. This is what makes the shared dump self-sufficient — see
"Orchestrating session" above for why, and "Excluding struck-through / grayed-out rows from review"
below for what the digest is for.

Getting that without paying per-cell COM costs needs a **three-level cascade**, because
`Font.Strikethrough`/`Font.Color` return `DBNull` when a range is mixed and a concrete value when it
is uniform:

1. Ask the whole `UsedRange` once. If it comes back "clean" (`Strikethrough = False` and a
   non-gray uniform `Color`), the sheet has nothing to strip — write the plain `Value2` dump and
   move on with **zero** extra COM calls. This is the common case for most sheets.
2. Otherwise ask each row once, over just that row's non-empty column span (3 COM calls per row).
   Rows that come back clean are skipped wholesale.
3. Only for rows that are struck or mixed, walk that row's non-empty cells; and only for cells that
   are themselves `DBNull` do the per-character `Characters(i,1)` reconstruction.

Measured on `SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx` (7 sheets, 1089-row worst case, ~2,000 struck and
~390 partially-struck cells): 65s end to end, well inside a 600s timeout. Raise the Bash/PowerShell
timeout for this step rather than skipping the cascade.

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
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class ExcelComWin32 {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
public static class XlsxDumpHelper {
    static string Cell(object[,] arr, object vals, int r, int c, int rows, int cols) {
        object v;
        if (arr == null) v = vals;
        else if (rows == 1) v = arr[1, c];
        else if (cols == 1) v = arr[r, 1];
        else v = arr[r, c];
        if (v == null) return null;
        return Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
    }
    // Non-empty column indices for one row — lets the scan walk only cells that have content.
    public static List<int> NonEmptyCols(object vals, int r, int rows, int cols) {
        var list = new List<int>();
        object[,] arr = vals as object[,];
        for (int c = 1; c <= cols; c++) {
            string s = Cell(arr, vals, r, c, rows, cols);
            if (s != null && s.Length > 0) list.Add(c);
        }
        return list;
    }
    // dead = coords to omit entirely; live = coords whose text is replaced by its unstruck remainder.
    // Both are keyed on ARRAY (UsedRange-relative) coordinates; r0/c0 shift the EMITTED coordinate
    // to sheet-absolute so a citation like [10,5] names the cell a reviewer sees in Excel.
    public static string FormatSheetLive(object vals, int rows, int cols,
                                         HashSet<long> dead, Dictionary<long,string> live,
                                         int r0, int c0) {
        var sb = new System.Text.StringBuilder();
        object[,] arr = vals as object[,];
        for (int r = 1; r <= rows; r++) {
            var parts = new List<string>();
            for (int c = 1; c <= cols; c++) {
                long key = ((long)r << 20) | (long)c;
                if (dead.Contains(key)) continue;
                string s = live.ContainsKey(key) ? live[key] : Cell(arr, vals, r, c, rows, cols);
                if (s != null && s.Length > 0) parts.Add("[" + (r + r0) + "," + (c + c0) + "]=" + s);
            }
            if (parts.Count > 0) { sb.Append(string.Join(" | ", parts)); sb.Append("\r\n"); }
        }
        return sb.ToString();
    }
}
"@

$out    = "<scratchpad dir>"
$path   = "<absolute path to the target workbook>"
$prefix = "<program id>"
# Sheet scoping. $null = every visible sheet. Set it when the user scopes the REV, e.g.
# @("表紙*","機能定義書*","帳票設計書*"). Hidden sheets (bk_ backups etc.) are out of scope by default.
$IncludeSheetPatterns = $null
$IncludeHiddenSheets  = $false

function Test-Gray([double]$argb) {
    $r = [int]$argb -band 0xFF; $g = ([int]$argb -shr 8) -band 0xFF; $b = ([int]$argb -shr 16) -band 0xFF
    return (($r -eq $g) -and ($g -eq $b) -and ($r -gt 80) -and ($r -lt 220))
}

# Record every EXCEL.EXE PID that already exists BEFORE launching our own instance. On this
# environment, `New-Object -ComObject Excel.Application` has been observed to sometimes attach to
# the user's own already-running interactive Excel process instead of spawning a genuinely new one
# (a Running-Object-Table quirk). If that happens and we later blindly call $excel.Quit(), it closes
# the user's ENTIRE real Excel session — including unrelated workbooks they had open — without
# saving. This bit us once already: a REV skill's dump step silently attached to the user's session
# and Quit() force-closed a workbook they had open for unrelated work, mid-dump.
$preExistingExcelPids = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
[uint32]$excelComPid = 0
[ExcelComWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelComPid) | Out-Null
if ($preExistingExcelPids -contains $excelComPid) {
    throw "Excel COM automation attached to the user's existing Excel process (PID $excelComPid) instead of creating a new instance. Aborting without calling Open/Close/Quit on it — ask the user to close their other Excel windows first."
}
$wb = $excel.Workbooks.Open($path, $true, $true)   # ReadOnly, no update-links prompt

$deleted       = New-Object System.Collections.Generic.List[object]
$skippedSheets = @()
$summary       = @()

foreach ($ws in $wb.Worksheets) {
    $n = $ws.Name
    if (-not $IncludeHiddenSheets -and $ws.Visible -ne -1) { continue }
    if ($IncludeSheetPatterns) {
        $hit = $false
        foreach ($pat in $IncludeSheetPatterns) { if ($n -like $pat) { $hit = $true; break } }
        if (-not $hit) { continue }
    }
    if ($n -like '詳細設計*' -or $n -like '*画面ｲﾒｰｼﾞ*') {
        $u = $ws.UsedRange
        $skippedSheets += "$n : rows=$($u.Rows.Count) cols=$($u.Columns.Count) (skipped — out of scope for every REV skill; not dumped)"
        continue
    }

    $used = $ws.UsedRange
    $rows = [Math]::Min($used.Rows.Count, 3000)
    $cols = [Math]::Min($used.Columns.Count, 220)
    $vals = $used.Value2
    # $vals is indexed from 1 WITHIN the UsedRange; $ws.Cells.Item is absolute on the sheet. Every
    # absolute access below adds this offset, and so does every coordinate written out. See
    # "UsedRange-relative vs sheet-absolute coordinates" below for why this matters.
    $r0 = $used.Row - 1
    $c0 = $used.Column - 1
    $dead    = New-Object 'System.Collections.Generic.HashSet[long]'
    $liveMap = New-Object 'System.Collections.Generic.Dictionary[long,string]'
    $nDead = 0; $nPart = 0

    # Level 1 — one question for the whole sheet. Clean sheets cost zero further COM calls.
    $wholeStrike = $used.Font.Strikethrough
    $wholeColor  = $used.Font.Color
    $sheetClean  = ($wholeStrike -isnot [System.DBNull]) -and ($wholeStrike -eq $false) -and
                   ($wholeColor  -isnot [System.DBNull]) -and (-not (Test-Gray $wholeColor))
    if (-not $sheetClean) {
        for ($r = 1; $r -le $rows; $r++) {
            $cl = [XlsxDumpHelper]::NonEmptyCols($vals, $r, $rows, $cols)
            if ($cl.Count -eq 0) { continue }
            # Level 2 — one question per row, over that row's non-empty span only.
            $ar = $r + $r0
            $rowRange = $ws.Range($ws.Cells.Item($ar, ($cl[0] + $c0)),
                                  $ws.Cells.Item($ar, ($cl[$cl.Count - 1] + $c0)))
            $rs = $rowRange.Font.Strikethrough
            $rc = $rowRange.Font.Color
            $drill = ($rs -is [System.DBNull]) -or ($rs -eq $true) -or
                     ($rc -is [System.DBNull]) -or (Test-Gray $rc)
            if (-not $drill) { continue }
            # Level 3 — per cell, and per character only where the cell itself is mixed.
            foreach ($c in $cl) {
                $cell = $ws.Cells.Item($ar, ($c + $c0))
                $s    = $cell.Font.Strikethrough
                $col  = $cell.Font.Color
                $isGray = (-not ($col -is [System.DBNull])) -and (Test-Gray $col)
                $key = (([int64]$r) -shl 20) -bor ([int64]$c)   # key stays array-relative
                $aC  = $c + $c0                                # digest records absolute coords
                if ($s -is [System.DBNull]) {
                    $raw = "$($cell.Value2)"; $lv = ""
                    for ($i = 1; $i -le $raw.Length; $i++) {
                        $ch = $cell.Characters($i, 1)
                        if (-not $ch.Font.Strikethrough) { $lv += $ch.Text }
                    }
                    if ($lv.Trim().Length -eq 0) {
                        [void]$dead.Add($key); $nDead++
                        $deleted.Add([PSCustomObject]@{S=$n; R=$ar; C=$aC; Kind='DEL'; Text=$raw})
                    } else {
                        $liveMap[$key] = $lv; $nPart++
                        $deleted.Add([PSCustomObject]@{S=$n; R=$ar; C=$aC; Kind='PART'; Text="raw='$raw' live='$lv'"})
                    }
                } elseif (($s -eq $true) -or $isGray) {
                    [void]$dead.Add($key); $nDead++
                    $kind = if ($s -eq $true) { 'DEL' } else { 'GRAY' }
                    $deleted.Add([PSCustomObject]@{S=$n; R=$ar; C=$aC; Kind=$kind; Text="$($cell.Value2)"})
                }
            }
        }
    }

    $safe = ($n -replace '[\\/:*?"<>|]','_')
    $text = [XlsxDumpHelper]::FormatSheetLive($vals, $rows, $cols, $dead, $liveMap, $r0, $c0)
    [System.IO.File]::WriteAllText((Join-Path $out ($prefix + "_" + $safe + ".txt")), $text, [System.Text.Encoding]::UTF8)
    $summary += "$n`t$($used.Rows.Count)x$($used.Columns.Count)`tdead=$nDead`tpartial=$nPart"
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# _DELETED_DIGEST.txt — removed content, grouped into contiguous row blocks. Structural filler is
# counted but not listed: a bare row/condition number, a hyphen placeholder, a comparison operator.
# Everything else is listed WHATEVER ITS LENGTH. Do not "simplify" this back to a character-count
# threshold — a length rule was tried first (drop anything under 4 chars) and silently swallowed 249
# of 761 omissions on SXJCB147, because this project's load-bearing tokens are routinely 1-3
# characters: block references like (5), footnote markers like ※2, query aliases Y/Z/A/C, and plain
# words such as 引数 / ﾗﾝｸ / 日付 / 参照先. Those are exactly what an asymmetry finding hangs on.
# The three filler classes below were verified safe to drop on that same workbook: all 248 numerics
# sat in No. columns (never in a 検索条件 right-hand side, so no literal value is lost), all 229
# hyphens and all 35 operators belonged to rows deleted in full, where the row's real content is
# already listed. Note this leaves ー (U+30FC, the katakana prolonged mark) out of the hyphen class
# on purpose — it is a letter, not a dash.
$lines = New-Object System.Collections.Generic.List[string]
foreach ($grp in $deleted | Group-Object S) {
    $items = $grp.Group | Sort-Object R, C
    $blockStart = $null; $prev = $null
    $buf = New-Object System.Collections.Generic.List[object]
    $flush = {
        if ($buf.Count -eq 0) { return }
        $subst = @($buf | Where-Object {
            $t = $_.Text.Trim()
            -not ($t -match '^[0-9]+$' -or $t -match '^[-‐‑–—―－]$' -or $t -match '^[=<>≠≦≧≤≥]+$')
        })
        $short = $buf.Count - $subst.Count
        $lines.Add("=== $($grp.Name) rows $blockStart-$prev ($($buf.Count) cells) ===")
        foreach ($it in $subst) { $lines.Add("  [$($it.R),$($it.C)] $($it.Kind): $($it.Text)") }
        if ($short -gt 0) { $lines.Add("  (+ $short filler cells omitted)") }
        $buf.Clear()
    }
    foreach ($it in $items) {
        if ($null -eq $blockStart) { $blockStart = $it.R }
        elseif ($it.R - $prev -gt 2) { & $flush; $blockStart = $it.R }
        $prev = $it.R; $buf.Add($it)
    }
    & $flush
}
[System.IO.File]::WriteAllText((Join-Path $out "_DELETED_DIGEST.txt"), ($lines -join "`r`n"), [System.Text.Encoding]::UTF8)

$summary -join "`n"
$skippedSheets -join "`n"   # a suspiciously large row count on a skipped 詳細設計書 sheet is a signal naming-standard-compliance may need to check its content directly
```

## UsedRange-relative vs sheet-absolute coordinates

`$used.Value2` is indexed from 1 **within the UsedRange**; `$ws.Cells.Item(r,c)` is absolute on the
sheet. Feeding array indices to `Cells.Item` reads a *different cell* on any sheet whose UsedRange
does not start at `A1`, and it fails silently. This script derives its non-empty column list from
the array, so before the fix every absolute access — the per-row `Range` for the level-2 scan and
the per-cell drill in level 3 — was off by the origin, and the emitted `[row,col]` tokens were off
too.

Both consequences matter, and the second is worse:

1. Strikethrough and gray-out get resolved against the wrong cells, so struck content can survive
   into the dump and live content can be dropped — the exact failure this live dump exists to
   prevent.
2. A finding citing `[10,5]` does not name the cell a reviewer opens. Reviewers navigate by these
   coordinates, so an offset dump makes every finding on that sheet unactionable.

**Measured on the 61 workbooks of `01_Doc\08_機能定義書\11_工程管理\PHASE3`**: 33 of 715 visible
multi-cell sheets do not start at `A1`; after this script's own `詳細設計*` / `*画面ｲﾒｰｼﾞ*` skips,
**12 of 600 in-scope sheets** are affected. Two are worth naming because they are sheets a REV
actually reads: `PXJCB134_流動停止(ﾊﾞｯﾁ).xlsx`'s `ﾒｰﾙｲﾒｰｼﾞ` (offset r1/c1 — the very sheet the skip
rule above carves out an exception *for*), and `SSJCB605_完成割合設定(ｻﾌﾞﾌﾟﾛ).xlsx`'s
`完成割合設定(ｻﾌﾞﾌﾟﾛ)_事業部説明なし` at offset **r33/c34**, where a cited `[5,3]` is really cell
`AK38`. `PXJCO124_ﾛｯﾄ振向け.xlsx`'s two `《参考》画面項目遷移` sheets (881 and 836 rows, r2/c0) and
`PXJCO192_処置指示登録.xlsx`'s `ﾛｯﾄ管理C　処置指示登録` (271x369, r0/c1) are the other large ones.

Reference masters are nearly clean by comparison — 1 of 84 visible sheets across the eight cached
registries and common-design files, and that one is `05.ｼｽﾃﾑ共通設計書.xlsx`'s
`(参考)配色ﾃﾝﾌﾟﾚｰﾄ`, which no check reads.

**The fix, applied in all three scripts here**: compute `$r0 = $used.Row - 1` / `$c0 = $used.Column - 1`
once per sheet, add it to every `Cells.Item` access, and add it to every coordinate written out. Keep
the internal `dead`/`live` keys array-relative — they index the same array — and offset only at the
boundary. `_DELETED_DIGEST.txt` records absolute coordinates too.

**This changes what cached dumps mean, and the source files did not change**, so mtime/length alone
would report a false cache hit. `meta.json` therefore carries `dumpFormat` (now `2`), compared
alongside mtime and length; a `dumpFormat 1` cache is treated as a miss and redumped. Nothing needs
to be deleted by hand.

### Parsing a dump line

Split the line on the literal `" | "` first, then take everything after `]=` **verbatim** — do not
trim the value. Cell values in these workbooks routinely carry significant leading or trailing
spaces (150 of the 980 non-empty cells on `PXJCO192_処置指示登録.xlsx`'s `ﾛｯﾄ管理C　処置指示登録`
sheet alone, including cells that are a single space used as a spacer), which makes a regex like
`\[(\d+),(\d+)\]=([^|]*)` plus a trim silently corrupt them. Verified: parsing verbatim reproduces
all 980 cells of that sheet exactly; the trimming version reported 149 false differences.

## Cross-session cache for reference/master files (not the target workbook)

The "dump once, share the text" section above only dedupes work *within* one REV run, for the one
target workbook under review. It does nothing for the other files a REV skill reads as ground
truth — `04.ﾒｯｾｰｼﾞ管理_*.xlsx`, `09.区分名称_step2.xlsx`,
`05.ｼｽﾃﾑ共通設計書.xlsx`, `06-*.xlsx` (機能一覧/DB一覧/レスポンス一覧), `07.共通項目取得.xlsx`, the
チェックリスト file, and every テーブルレイアウト workbook. These almost never change between one
program's REV and the next, yet without caching they get re-opened via Excel COM from scratch every
single time, by every agent that needs them — including more than once within the same REV run, when
several sibling skills happen to need the same reference file.

**One exception: `82.画面項目辞書_*.xlsx` is indexed, not cached — see
`_shared/reference-index.md`.** Caching still re-reads whole sheets into an agent's context, and
the only thing any check wants from that file is `id → 画面項目名`. On the 工程管理 copy, the two
sheets a check reads dump to roughly 4.2M characters (~1,318,000 tokens) against a ~56k-character
(~28,200-token) index, or ~1,100 tokens once subset to one program's IDs — **exact figures and the
source state they were measured at live in `reference-index.md`, not here**, so they only have to be
refreshed in one place. Every other file listed above — including テーブルレイアウト workbooks and the
`06-*.xlsx` registries — keeps using the cache below; that doc explains why each of those still
needs a builder of its own before it can be indexed safely.

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
- **`04.ﾒｯｾｰｼﾞ管理_<WG>.xlsx`** — confirmed for real on the 工程管理 copy: `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx`
  has 11 sheets (改訂履歴, 方針, XSA(共通), XSK(共通), XJZ(共通), XJA(基準), XJB(受注), **XJC(工程)**,
  **SJC(工程)**, XJD(品質), XJE(生産)) but the skills' own routing rules (see
  `design-doc-internal-consistency` check 3/4) mean only the reviewed program's own-WG JOBコード
  sheets (bold above) are ever read from THIS file — a common (共通) prefix routes to
  `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` instead, and a foreign-WG prefix routes to that other WG's own file, never
  to a same-named sheet embedded in this one. Use
  `$OnlySheetPatterns = @("XJC(*", "SJC(*")` (or the equivalent for the WG you're checking) —
  `-like` wildcard patterns, not exact names, since a WG's own-prefix sheet can carry an extra
  suffix that an exact-name match would miss. **Leave the pattern open-ended with no closing `)`** —
  `-like` requires a literal `)` to match the actual end of the string, so a closed pattern like
  `"SJC(*)"` fails against a name such as `SJC(工程)再開発追加分 ` (extra text AND a trailing space,
  so it doesn't end in `)`) even though it very much should match; confirmed this exact bug while
  testing against the real 画面項目辞書 file, which carries a sheet named exactly that.

**`82.画面項目辞書_*.xlsx` is not on this list** — it is indexed rather than dumped, so it never
reaches `$OnlySheetPatterns` at all. Its sheet selection is the index builder's business and works
by header detection, not by name; see `_shared/reference-index.md`.

Leave `$OnlySheetPatterns` `$null` (the default) for file types with no such known-safe restriction
(`09.区分名称_step2.xlsx`, `05.ｼｽﾃﾑ共通設計書.xlsx`, `06-*.xlsx`, `07.共通項目取得.xlsx`, the
checklist file, `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` — these either have few sheets
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

$DumpFormat = 2       # 1 = coordinates relative to the UsedRange; 2 = sheet-absolute

$cacheValid = $false
if (Test-Path $metaPath) {
    try {
        $meta = Get-Content -Raw -Encoding UTF8 $metaPath | ConvertFrom-Json
        # dumpFormat guards against a cache written by an older script whose coordinates mean
        # something different. The source file is unchanged in that case, so mtime/length alone
        # would wrongly report a hit.
        if ($meta.sourcePath -eq $SourcePath -and $meta.lastWriteTimeUtc -eq $currentMTime -and
            $meta.length -eq $currentLength -and $meta.dumpFormat -eq $DumpFormat) {
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
    // r0/c0 = UsedRange origin offset, so emitted coordinates are sheet-absolute.
    public static string FormatSheet(object vals, int rows, int cols, int r0, int c0) {
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
                    if (s.Length > 0) parts.Add("[" + (r + r0) + "," + (c + c0) + "]=" + s);
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
        $text = [XlsxDumpHelperCache]::FormatSheet($vals, $rows, $cols,
                                                   ($used.Row - 1), ($used.Column - 1))
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
        dumpFormat = $DumpFormat
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
$DumpFormat = 2       # 1 = coordinates relative to the UsedRange; 2 = sheet-absolute

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
            if ($meta.sourcePath -eq $info.SourcePath -and $meta.lastWriteTimeUtc -eq $info.CurrentMTime -and
                $meta.length -eq $info.CurrentLength -and $meta.dumpFormat -eq $DumpFormat) { $valid = $true }
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
    // r0/c0 = UsedRange origin offset, so emitted coordinates are sheet-absolute.
    public static string FormatSheet(object vals, int rows, int cols, int r0, int c0) {
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
                    if (s.Length > 0) parts.Add("[" + (r + r0) + "," + (c + c0) + "]=" + s);
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
            $text = [XlsxDumpHelperBatch]::FormatSheet($vals, $rows, $cols,
                                                       ($used.Row - 1), ($used.Column - 1))
            [System.IO.File]::WriteAllText($file, $text, [System.Text.Encoding]::UTF8)
            Write-Output "$($ws.Name)`t$file"
        }
        $wb.Close($false)

        @{
            sourcePath = $info.SourcePath; lastWriteTimeUtc = $info.CurrentMTime; length = $info.CurrentLength
            onlySheetPatterns = $info.OnlySheetPatterns; dumpedSheets = $dumpedSheetNames
            dumpedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
            dumpFormat = $DumpFormat
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

**This is already done for you. Do not run your own strikethrough or gray-out scan.** The dump
script resolves both while the workbook is open: fully-struck and gray cells are absent from the
`.txt`, and partially-struck cells carry only their live text. A row you can see in the dump is
live; a row that was struck simply is not there. Re-deriving this per agent was the single largest
source of duplicated cost (~49k tokens × 5-6 agents) **and** of false findings — see "What this
prevents" below.

Read this section to understand what the dump has already decided on your behalf, and when to reach
for `_DELETED_DIGEST.txt`.

### Why the content is excluded at all

Design docs are often copied from an older workbook and edited in place. This project's writing-rule
checklist (`01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx`, sheet
"ﾁｪｯｸﾘｽﾄ", item 1-5) says authors should delete struck-through rows entirely and turn red
"to-be-fixed" text back to black when reusing a workbook this way — but that cleanup is often
skipped, leaving struck-through/grayed rows still physically present. **Content marked this way is
inactive/deleted-in-spirit and must never be treated as something the program actually
references** — don't count it toward a "referenced columns" list, don't flag it as a
naming/consistency violation, don't cite it as evidence of what the current design does. Note that
red alone does *not* mean deleted (red is the "to-be-fixed" marker and is still live); only
strikethrough and gray are.

A one-line "N struck-through/grayed cells excluded at dump time" note in the final report is enough —
never enumerate them as findings. The per-sheet `dead=`/`partial=` counts the dump script prints are
exactly that number.

**An empty or near-empty live dump for a sheet is meaningful, not a dump failure.** When an entire
更新条件表/画面設計書 sheet comes back struck top to bottom — including its own header — the correct
read is that the **whole sheet is superseded** (commonly: an old JAGUR-era 更新条件表(<TableID>)
sheet that a newer 更新条件表(<TableID>WF) sheet has replaced). Exclude the whole thing from the
review and say so. Per explicit user direction, red+strikethrough always means "exclude from REV
scope" in this project, with no exception for how broadly it is applied — do not reason your way out
of it ("this can't be meaningful, it's on every row including the header, must be leftover noise").
That reasoning has been tried and explicitly overruled.

### When to read `_DELETED_DIGEST.txt`

The digest holds what was removed, grouped into contiguous row blocks, with `DEL` (whole cell),
`GRAY` (whole cell, gray) and `PART` (`raw=` vs `live=` for a partially-struck cell) entries. Only
structural filler is left out — a bare row/condition number, a hyphen placeholder, a comparison
operator — and each block still reports how many it dropped (`+ N filler cells omitted`). Short but
meaningful cells are listed in full: `(5)`, `※2`, an alias `Y`, `引数`, `ﾗﾝｸ` and the like all
survive the filter, because deletion asymmetries usually hang on exactly those.

Most checks never need it — they are asking "is what the doc *currently* says correct?", and the live
dump answers that directly. Reach for the digest only when a finding turns on **what was deleted**,
which in practice is `design-doc-internal-consistency`:

- A section whose counterpart was deleted while it stayed live — e.g. a ※-note whose only reference
  was removed, or a "(通常/識別ｶｰﾄﾞ)" wording left behind after the 識別ｶｰﾄﾞ output rows were deleted.
  The live dump shows the survivor; only the digest shows that its partner is gone.
- A step-number or branch table with a hole in it, where the missing branch was struck rather than
  renumbered.
- A "削除" revision note sitting next to content that is *not* struck — a real and load-bearing
  ambiguity (confirmed on `SXJCB147`'s `帳票設計書(RXJC042)` 検索条件 No.5, where the surrounding
  same-day deletions *were* struck and this one was not).

If you are about to report a section as "a leftover that should have been deleted but wasn't cleaned
up," check the digest first — it may already be marked dead, in which case there is nothing to fix.
Confirmed for real on `XJC_ｼｽﾃﾑ共通設計書.xlsx`, sheet `ﾛｯﾄ停止ﾁｪｯｸ`, rows 829-832 (a
"結合条件(A INNER JOIN B)" block left over after its alias B was reassigned from a real table to a
delegated get-item sub-block): a review reported this as a live, not-yet-cleaned-up defect, but every
cell in the block was `Font.Strikethrough=True` — it had already been correctly deleted, and the
"defect" was a pure false positive.

### What this prevents

Every one of these was a real false finding produced by an agent doing its own scan, or skipping it:

- **Six tables reported as "used but not declared in the I/O table"** (PXJCO193), entirely because
  the scan was skipped on a 1000+ row sheet whose 参照ｴﾝﾃｨﾃｨ blocks were red-and-struck-through 100+
  rows deep. Sheet size was exactly why the scan mattered most, and exactly why it got shortcut.
- **`帳票設計書(RSJC035)`'s `注意事項備考` reported as a column that doesn't exist in `TXJCM137`**
  (`SXJCB147`, cells `F202`/`Q202`): raw text was `注意事項備考` / `A.注意事項備考`, but only
  `注意事項` was live. TXJCM137 has `備考` and `注意事項` as separate columns — the doc was correctly
  referencing the latter with a stale, not-fully-deleted `備考` glued on.
- **`COUNT(A.層数層No)` reported as referencing a nonexistent column** and "fixed" to
  `COUNT(A.層No)` (`PXJCO124_ﾛｯﾄ振向け.xlsx`, `ﾁｪｯｸ処理設計書(GXJC124A)`, `[280,17]`) — `層No` was
  already the live text; `層数` was the struck leftover. The doc was already correct.
- **`"GSXJC205A"` reported as a typo of the screen ID `GSJC205A`** (`PSJCO205_返品処置指示発行.xlsx`,
  `画面設計書(GSJC205A)`, `[333,28]`) — only the `X` was struck; live text was an exact match.
- **`③④` treated as two live entity references** (`PSJCO205`, `更新条件表(TXJCM006)`, rows 74/84/95,
  col 16 取得元) when only `④` was live — the sheet's own row-9 legend documents the renumbering that
  stranded `③`.

The last four share one shape: a **mixed-formatting cell**, where `Font.Strikethrough` returns
`System.DBNull` rather than `True`/`False`. That is why the dump script drops to per-character
reconstruction on exactly those cells. The trap is that `System.DBNull` interpolates into a
PowerShell string as an *empty string* — `Write-Output "strike=$whole"` prints `strike=` whether the
value is `False` or `DBNull` — so an ad-hoc scan can read "mixed" as "not struck, live" and never
notice. Test the type explicitly (`-is [System.DBNull]`) and never eyeball a printed value. The dump
script does this; a hand-rolled scan usually doesn't.

### One judgment call the dump can't make for you

**Parallel `処理区分="..."の場合` (or similarly-named) branch variants inside one sheet.** A common
shape in this project's 画面設計書 is two or more parallel sub-blocks handling different values of the
same discriminator (e.g. `品目コード` vs `管理No`, one per branch), often followed by several more
sub-blocks that logically belong to just one of those branches. When one branch is deprecated, the
strikethrough frequently covers not just that one block but every subsequent block that depended on
it too, spanning hundreds of rows and several numbered sub-sections in a row.

In the live dump this appears as a *gap* — a discriminator with no handling. Whether that gap is a
correct deprecation or a genuine design hole is the finding, and it needs the digest plus your own
reading of the surrounding branch structure. `SXJCB147`'s 部門GRP="SC200"(SMD) row, left with no
帳票 after RSJC034/RSJC036 were deleted, is the canonical example of the "genuine hole" case.

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
