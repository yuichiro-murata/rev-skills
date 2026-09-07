# Detecting font-size and cell-merge irregularities

This is specific to `design-doc-formatting-consistency` — no other REV skill needs this technique,
so it lives in its own file rather than in `_shared/xlsx-excel-com-dump.md` (which every REV skill
reads) to avoid making every other skill pay the token cost of a section it never uses.

A separate class of defect from strikethrough: a cell whose **font size** or **merge span** breaks
from the surrounding pattern, usually left over from a copy-paste edit that didn't fully take (a
pasted row keeping the source's font size, a header block duplicated on the sheet's right side that
never got updated when the left side did). Both are checkable via COM the same way as
Strikethrough/Color, and both need the same DBNull guard for cells with mixed-formatting runs
(`if ($cell.Font.Size -is [System.DBNull]) { ... }` — treat as "mixed, inspect manually" rather than
crashing the scan).

**Font size**: loop the sheet's non-empty cells (reuse the coordinates from the bulk `Value2` dump,
same as the strikethrough scan — no need to touch empty cells), read `Font.Size` per cell, tally into
a per-sheet frequency count, and take the most common value as that sheet's "mode." Cells whose size
differs from the mode are candidates — but **most of what this turns up is a deliberate, consistent
convention, not a defect**, and reporting all of it is far too noisy to be useful. Before surfacing
anything, recognize and silently exclude these recurring systematic patterns (seen consistently
across every workbook reviewed this way so far):
- A dedicated "改訂履歴" annotation column (commonly column 105/106 in this project's templates)
  deliberately set smaller than body text, appearing on potentially hundreds of rows.
- The sheet-title/section-header template cells every sheet shares (row 1's "2C"/"3C" markers, the
  sheet-name label, "Ⅰ．"–"Ⅷ．" section headings) — a fixed, larger size across all workbooks.
- Column-group headers in a "Ⅵ．画面項目制御" -style matrix (e.g. `処理区分="..."の場合` column
  labels) — deliberately smaller, consistently so across every instance of that matrix in the sheet.
- Any 画面イメージ/mockup-screenshot sheet — a UI mockup naturally mixes title/label/icon sizes; this
  is not something to flag at all.
Only report a font-size deviation that's **genuinely isolated** — a single cell or a small handful
that don't fit any repeating pattern above, especially ones sitting inside an otherwise-uniform
repeating list (e.g. two cells in one 更新条件表's item list are smaller than every other row in that
same list) — those are the ones that plausibly indicate an accidental leftover format from editing.

**Cell merges**: for the same non-empty cells, check `cell.MergeCells`; if true, only record it when
the cell is the merge's own anchor (`cell.MergeArea.Row -eq r -and .Column -eq c`, since a merged
range's other cells all report `MergeCells=True` too but aren't independently meaningful), and note
its `MergeArea.Rows.Count`/`.Columns.Count` (the span). **Do not compare merge spans sheet-wide by
column position** the way font-size is compared against a sheet-wide mode — a single column position
is legitimately reused by many unrelated sections with genuinely different natural spans (a title row,
a data row, a notes row), so a global "most common span at this column" produces overwhelming noise
(tried this first; it surfaced tens of thousands of false positives). Instead, compare **locally**:
sort each column's merge records by row, and flag a pair of same-column merges whose spans differ
while sitting within a few rows of each other (roughly 4 or fewer) — close enough to plausibly belong
to the same repeating list or the same duplicated header block, where a span mismatch is far more
likely to be a genuine inconsistency (e.g. a duplicated header block on the sheet's right side that
kept stale content/labels, or one row of an item list merged differently than its siblings) rather
than two unrelated sections that just happen to share a column index. This still needs a manual pass
to drop template-boilerplate coincidences (e.g. a "No." column header naturally spanning differently
right before a "特記事項" footer row) — report only the pairs that look like an actual duplicated
block or list, not just adjacent structural cells of different kinds.
