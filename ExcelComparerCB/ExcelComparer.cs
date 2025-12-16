using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;

namespace ExcelComparerCB;

public sealed class ExcelComparer
{
    public async Task<ComparisonResult> CompareAsync(
        string fileA,
        string fileB,
        ComparisonOptions options,
        IProgress<ProgressInfo>? progress,
        CancellationToken ct)
    {
        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();

            var result = new ComparisonResult();
            progress?.Report(new ProgressInfo(1, "Leyendo estructura de libros..."));

            using var docA = SpreadsheetDocument.Open(fileA, false);
            using var docB = SpreadsheetDocument.Open(fileB, false);

            var wbA = ReadWorkbook(docA);
            var wbB = ReadWorkbook(docB);

            progress?.Report(new ProgressInfo(5, "Comparando hojas..."));
            DiffSheets(wbA, wbB, result, options);

            if (options.CompareSheetOrder)
            {
                DiffSheetOrder(docA, docB, result);
            }

            // Build number-format caches once per workbook to avoid repeated stylesheet scans
            var nfCacheA = BuildNumberFormatCache(docA.WorkbookPart!);
            var nfCacheB = BuildNumberFormatCache(docB.WorkbookPart!);

            var allSheetNames = wbA.SheetsByName.Keys
                .Union(wbB.SheetsByName.Keys)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            int total = allSheetNames.Count;
            for (int i = 0; i < total; i++)
            {
                ct.ThrowIfCancellationRequested();

                var sheetName = allSheetNames[i];
                var pct = 10 + (int)((i / (double)Math.Max(1, total)) * 85);

                // Throttle progress updates slightly by reporting per sheet only
                progress?.Report(new ProgressInfo(pct, $"Comparando hoja: {sheetName} ({i + 1}/{total})"));

                if (!wbA.SheetsByName.TryGetValue(sheetName, out var sA) ||
                    !wbB.SheetsByName.TryGetValue(sheetName, out var sB))
                {
                    continue;
                }

                if (!options.IncludeHiddenSheets && (sA.Hidden || sB.Hidden))
                    continue;

                // Retrieve WorksheetPart once per sheet and materialize Worksheet once to avoid repeated cost
                var wbPartA = docA.WorkbookPart!;
                var wbPartB = docB.WorkbookPart!;
                var wsPartA = (WorksheetPart)wbPartA.GetPartById(sA.RelId);
                var wsPartB = (WorksheetPart)wbPartB.GetPartById(sB.RelId);
                var wsA = wsPartA.Worksheet;
                var wsB = wsPartB.Worksheet;

                if (options.CompareUsedRange)
                    DiffUsedRange(wsA, wsB, sheetName, result);

                if (options.CompareValidations)
                    DiffDataValidations(wsA, wsB, sheetName, result);

                if (options.CompareConditionalFormats)
                    DiffConditionalFormatting(wsA, wsB, sheetName, result);

                if (options.CompareHiddenRowsCols)
                    DiffHiddenRowsCols(wsA, wsB, sheetName, result);

                var cellsA = ReadCells(wbPartA, wsA, wbPartA.SharedStringTablePart?.SharedStringTable, options, nfCacheA);
                var cellsB = ReadCells(wbPartB, wsB, wbPartB.SharedStringTablePart?.SharedStringTable, options, nfCacheB);

                DiffCells(sheetName, cellsA, cellsB, result, options, wbPartA, wbPartB);
            }

            progress?.Report(new ProgressInfo(100, "Finalizado."));
            return result;

        }, ct);
    }

    // -------- Workbook reading --------

    private sealed class WorkbookInfo
    {
        public Dictionary<string, SheetInfo> SheetsByName { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> SheetOrder { get; } = new();
    }

    private sealed class SheetInfo
    {
        public required string Name { get; init; }
        public required string SheetId { get; init; }
        public required string RelId { get; init; }
        public bool Hidden { get; init; }
        public bool VeryHidden { get; init; }
    }

    private static WorkbookInfo ReadWorkbook(SpreadsheetDocument doc)
    {
        var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart no encontrado.");
        var wb = wbPart.Workbook ?? throw new InvalidOperationException("Workbook no encontrado.");

        var info = new WorkbookInfo();

        foreach (var s in wb.Sheets!.OfType<Sheet>())
        {
            var name = s.Name?.Value ?? "(sin nombre)";
            var state = s.State?.Value; // Visible / Hidden / VeryHidden

            info.SheetsByName[name] = new SheetInfo
            {
                Name = name,
                SheetId = s.SheetId?.Value.ToString(CultureInfo.InvariantCulture) ?? "",
                RelId = s.Id?.Value ?? "",
                Hidden = state == SheetStateValues.Hidden,
                VeryHidden = state == SheetStateValues.VeryHidden
            };
            info.SheetOrder.Add(name);
        }

        return info;
    }

    private static void DiffSheets(WorkbookInfo a, WorkbookInfo b, ComparisonResult result, ComparisonOptions options)
    {
        var names = a.SheetsByName.Keys.Union(b.SheetsByName.Keys, StringComparer.OrdinalIgnoreCase);

        foreach (var name in names.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            var hasA = a.SheetsByName.TryGetValue(name, out var sA);
            var hasB = b.SheetsByName.TryGetValue(name, out var sB);

            if (hasA && !hasB)
            {
                result.Diffs.Add(new DiffItem(name, "", DiffKind.Removed, "Sheet", "Present", "Missing"));
                continue;
            }
            if (!hasA && hasB)
            {
                result.Diffs.Add(new DiffItem(name, "", DiffKind.Added, "Sheet", "Missing", "Present"));
                continue;
            }

            // ambos
            if (sA!.Hidden != sB!.Hidden || sA.VeryHidden != sB.VeryHidden)
            {
                var before = sA.VeryHidden ? "VeryHidden" : (sA.Hidden ? "Hidden" : "Visible");
                var after = sB!.VeryHidden ? "VeryHidden" : (sB.Hidden ? "Hidden" : "Visible");
                result.Diffs.Add(new DiffItem(name, "", DiffKind.Modified, "SheetVisibility", before, after));
            }
        }
    }

    private static void DiffSheetOrder(SpreadsheetDocument docA, SpreadsheetDocument docB, ComparisonResult result)
    {
        var orderA = docA.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().Select(s => s.Name!.Value!).ToList();
        var orderB = docB.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().Select(s => s.Name!.Value!).ToList();

        // Report per-sheet position changes for better traceability
        var indexA = orderA.Select((name, idx) => (name, idx)).ToDictionary(t => t.name, t => t.idx, StringComparer.OrdinalIgnoreCase);
        var indexB = orderB.Select((name, idx) => (name, idx)).ToDictionary(t => t.name, t => t.idx, StringComparer.OrdinalIgnoreCase);
        foreach (var name in indexA.Keys.Intersect(indexB.Keys, StringComparer.OrdinalIgnoreCase))
        {
            var ia = indexA[name];
            var ib = indexB[name];
            if (ia != ib)
            {
                result.Diffs.Add(new DiffItem(name, "", DiffKind.Modified, "SheetOrderIndex", ia.ToString(CultureInfo.InvariantCulture), ib.ToString(CultureInfo.InvariantCulture)));
            }
        }
    }

    // -------- Worksheet-level diffs --------

    private static void DiffUsedRange(Worksheet wsA, Worksheet wsB, string sheetName, ComparisonResult result)
    {
        var urA = GetUsedRangeFromWorksheet(wsA);
        var urB = GetUsedRangeFromWorksheet(wsB);
        if (urA != urB)
        {
            result.Diffs.Add(new DiffItem(sheetName, "", DiffKind.Modified, "UsedRange", urA, urB));
        }
    }

    private static string GetUsedRangeFromWorksheet(Worksheet ws)
    {
        var dim = ws.SheetDimension?.Reference?.Value;
        if (!string.IsNullOrWhiteSpace(dim)) return dim!;

        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
        if (sheetData is null) return "";

        int maxRow = 0;
        int maxCol = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            if (row.RowIndex != null && row.RowIndex.Value > maxRow) maxRow = (int)row.RowIndex.Value;
            // iterate cells of row only
            foreach (var cell in row.Elements<Cell>())
            {
                var addr = cell.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(addr)) continue;
                ParseAddress(addr, out int r, out int c);
                if (r > maxRow) maxRow = r;
                if (c > maxCol) maxCol = c;
            }
        }
        if (maxRow == 0 || maxCol == 0) return "";
        return $"A1:{ColumnIndexToName(maxCol)}{maxRow}";
    }

    private static void DiffDataValidations(Worksheet wsA, Worksheet wsB, string sheetName, ComparisonResult result)
    {
        var valsA = GetDataValidationSummaryFromWorksheet(wsA);
        var valsB = GetDataValidationSummaryFromWorksheet(wsB);
        if (!StringEquals(valsA, valsB))
        {
            result.Diffs.Add(new DiffItem(sheetName, "", DiffKind.Modified, "DataValidation", valsA, valsB));
        }
    }

    private static string GetDataValidationSummaryFromWorksheet(Worksheet ws)
    {
        var dv = ws.Descendants<DataValidation>().ToList();
        if (dv.Count == 0) return "";
        var sb = new StringBuilder();
        foreach (var d in dv)
        {
            var sqref = d.SequenceOfReferences?.InnerText ?? "";
            var type = d.Type?.Value.ToString() ?? "";
            var op = d.Operator?.Value.ToString() ?? "";
            var f1 = d.Formula1?.Text ?? "";
            var f2 = d.Formula2?.Text ?? "";
            sb.Append('[').Append(sqref).Append("; ").Append(type).Append(' ').Append(op).Append("; ")
              .Append(f1).Append(' ').Append(f2).Append(']');
        }
        return sb.ToString();
    }

    private static void DiffConditionalFormatting(Worksheet wsA, Worksheet wsB, string sheetName, ComparisonResult result)
    {
        var cfA = wsA.Descendants<ConditionalFormatting>().Select(cf => cf.InnerText).ToList();
        var cfB = wsB.Descendants<ConditionalFormatting>().Select(cf => cf.InnerText).ToList();
        if (!cfA.SequenceEqual(cfB))
        {
            result.Diffs.Add(new DiffItem(sheetName, "", DiffKind.Modified, "ConditionalFormatting",
                string.Join("|", cfA), string.Join("|", cfB)));
        }
    }

    private static void DiffHiddenRowsCols(Worksheet wsA, Worksheet wsB, string sheetName, ComparisonResult result)
    {
        var colsA = wsA.Elements<Columns>().FirstOrDefault()?.Elements<Column>().Where(c => c.Hidden != null && c.Hidden.Value).Select(c => $"{c.Min}-{c.Max}").OrderBy(x => x).ToList() ?? new();
        var colsB = wsB.Elements<Columns>().FirstOrDefault()?.Elements<Column>().Where(c => c.Hidden != null && c.Hidden.Value).Select(c => $"{c.Min}-{c.Max}").OrderBy(x => x).ToList() ?? new();
        if (!colsA.SequenceEqual(colsB))
            result.Diffs.Add(new DiffItem(sheetName, "", DiffKind.Modified, "HiddenColumns", string.Join(",", colsA), string.Join(",", colsB)));

        var rowsA = wsA.Descendants<Row>().Where(r => r.Hidden != null && r.Hidden.Value).Select(r => r.RowIndex!.Value).OrderBy(x => x).ToList();
        var rowsB = wsB.Descendants<Row>().Where(r => r.Hidden != null && r.Hidden.Value).Select(r => r.RowIndex!.Value).OrderBy(x => x).ToList();
        if (!rowsA.SequenceEqual(rowsB))
            result.Diffs.Add(new DiffItem(sheetName, "", DiffKind.Modified, "HiddenRows", string.Join(",", rowsA), string.Join(",", rowsB)));
    }

    // -------- Cell reading + diff --------

    private sealed class CellInfo
    {
        public string? ValueText { get; init; }     // normalizado a texto
        public string? FormulaText { get; init; }   // fórmula tal cual
        public uint? StyleIndex { get; init; }      // styleId
        public string? NumberFormatCode { get; init; } // resolved from styles
    }

    // New ReadCells variant that accepts a materialized Worksheet and a prebuilt number-format cache
    private static Dictionary<string, CellInfo> ReadCells(WorkbookPart wbPart, Worksheet ws, SharedStringTable? sst, ComparisonOptions options, Dictionary<int, string?>? nfCache)
    {
        var sheetData = ws.Elements<SheetData>().FirstOrDefault();

        var dict = new Dictionary<string, CellInfo>(StringComparer.OrdinalIgnoreCase);
        if (sheetData is null) return dict;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var addr = cell.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(addr)) continue;

                var formula = cell.CellFormula?.Text;
                var val = ReadCellValueAsText(cell, sst);

                uint? styleIdx = null;
                string? nfCode = null;
                if (options.CompareCellFormat)
                {
                    styleIdx = cell.StyleIndex?.Value;
                    if (styleIdx != null && nfCache != null && nfCache.TryGetValue((int)styleIdx.Value, out var code))
                    {
                        nfCode = code;
                    }
                    else
                    {
                        nfCode = ResolveNumberFormatCode(wbPart, styleIdx);
                    }
                }

                if (val is null && formula is null && !options.CompareCellFormat) continue;

                var ci = new CellInfo
                {
                    ValueText = val,
                    FormulaText = formula,
                    StyleIndex = styleIdx,
                    NumberFormatCode = nfCode
                };

                dict[addr] = ci;
            }
        }

        return dict;
    }

    private static string? ReadCellValueAsText(Cell cell, SharedStringTable? sst)
    {
        if (cell.CellValue is null)
        {
            // Some inline strings use InlineString
            if (cell.DataType?.Value == CellValues.InlineString)
            {
                return cell.InlineString?.Text?.Text ?? cell.InlineString?.InnerText;
            }
            return null;
        }
        var raw = cell.CellValue.Text;

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (sst is null) return raw;
            if (!int.TryParse(raw, out var idx)) return raw;

            var item = sst.Elements<SharedStringItem>().ElementAtOrDefault(idx);
            return item?.InnerText ?? raw;
        }

        return raw;
    }

    private static Dictionary<int, string?> BuildNumberFormatCache(WorkbookPart wbPart)
    {
        var result = new Dictionary<int, string?>();
        var styles = wbPart.WorkbookStylesPart?.Stylesheet;
        if (styles is null) return result;

        var cellXfs = styles.CellFormats?.Elements<CellFormat>().ToList();
        if (cellXfs is null) return result;

        var numbering = styles.NumberingFormats?.Elements<NumberingFormat>().Where(n => n.NumberFormatId != null).ToDictionary(n => (int)n.NumberFormatId!.Value, n => n.FormatCode?.Value) ?? new Dictionary<int, string?>();

        for (int i = 0; i < cellXfs.Count; i++)
        {
            var xf = cellXfs[i];
            if (xf.NumberFormatId == null)
            {
                result[i] = null;
                continue;
            }
            var nfid = (int)xf.NumberFormatId.Value;
            if (numbering.TryGetValue(nfid, out var code)) result[i] = code;
            else result[i] = nfid.ToString(CultureInfo.InvariantCulture);
        }

        return result;
    }

    private static string? ResolveNumberFormatCode(WorkbookPart wbPart, uint? styleIndex)
    {
        if (styleIndex is null) return null;
        var styles = wbPart.WorkbookStylesPart?.Stylesheet;
        if (styles is null) return null;
        var cellXfs = styles.CellFormats?.Elements<CellFormat>().ToList();
        if (cellXfs is null) return null;
        var idx = (int)styleIndex.Value;
        if (idx < 0 || idx >= cellXfs.Count) return null;
        var xf = cellXfs[idx];
        if (xf.NumberFormatId == null) return null;
        var nfid = (int)xf.NumberFormatId.Value;
        // Try custom number formats
        var nfs = styles.NumberingFormats?.Elements<NumberingFormat>().FirstOrDefault(n => n.NumberFormatId != null && n.NumberFormatId.Value == nfid);
        if (nfs != null) return nfs.FormatCode?.Value;
        // Built-in formats: we can return the id
        return nfid.ToString(CultureInfo.InvariantCulture);
    }

    private static void DiffCells(
        string sheetName,
        Dictionary<string, CellInfo> a,
        Dictionary<string, CellInfo> b,
        ComparisonResult result,
        ComparisonOptions options,
        WorkbookPart wbA,
        WorkbookPart wbB)
    {
        // Avoid sorting large key sets. Use two-pass iteration with a HashSet to de-duplicate in O(n).
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sheetDiffs = new List<DiffItem>();
        sheetDiffs.Capacity = Math.Max( (a?.Count ?? 0) + (b?.Count ?? 0), 0);

        void ProcessKey(string addr)
        {
            if (!seen.Add(addr)) return;

            var hasA = a.TryGetValue(addr, out var cA);
            var hasB = b.TryGetValue(addr, out var cB);

            if (hasA && !hasB)
            {
                sheetDiffs.Add(new DiffItem(sheetName, addr, DiffKind.Removed, "Cell", Summarize(cA, options), null));
                return;
            }
            if (!hasA && hasB)
            {
                sheetDiffs.Add(new DiffItem(sheetName, addr, DiffKind.Added, "Cell", null, Summarize(cB, options)));
                return;
            }

            var changes = new List<DiffItem>();

            if (options.CompareValues)
            {
                var vA = cA!.ValueText;
                var vB = cB!.ValueText;
                if (!StringEquals(vA, vB))
                    changes.Add(new DiffItem(sheetName, addr, DiffKind.Modified, "Value", vA, vB));
            }

            if (options.CompareFormulas)
            {
                var fA = cA!.FormulaText;
                var fB = cB!.FormulaText;
                if (!StringEquals(fA, fB))
                    changes.Add(new DiffItem(sheetName, addr, DiffKind.Modified, "Formula", fA, fB));
            }

            if (options.CompareCellFormat)
            {
                var sA = cA!.StyleIndex?.ToString();
                var sB = cB!.StyleIndex?.ToString();
                if (!StringEquals(sA, sB))
                    changes.Add(new DiffItem(sheetName, addr, DiffKind.Modified, "StyleIndex", sA, sB));

                var nfA = cA!.NumberFormatCode;
                var nfB = cB!.NumberFormatCode;
                if (!StringEquals(nfA, nfB))
                    changes.Add(new DiffItem(sheetName, addr, DiffKind.Modified, "NumberFormat", nfA, nfB));
            }

            foreach (var ch in changes)
                sheetDiffs.Add(ch);
        }

        if (a != null)
        {
            foreach (var k in a.Keys)
                ProcessKey(k);
        }
        if (b != null)
        {
            foreach (var k in b.Keys)
                ProcessKey(k);
        }

        // Batch add: reduces pressure if UI is observing the list
        if (sheetDiffs.Count > 0)
            result.Diffs.AddRange(sheetDiffs);
    }

    private static string? Summarize(CellInfo? c, ComparisonOptions options)
    {
        if (c is null) return null;
        var sb = new StringBuilder();
        if (options.CompareFormulas && !string.IsNullOrWhiteSpace(c.FormulaText))
            sb.Append("=").Append(c.FormulaText);

        if (options.CompareValues && !string.IsNullOrWhiteSpace(c.ValueText))
        {
            if (sb.Length > 0) sb.Append(" | ");
            sb.Append(c.ValueText);
        }
        if (options.CompareCellFormat)
        {
            if (sb.Length > 0) sb.Append(" | ");
            sb.Append("style:").Append(c.StyleIndex?.ToString() ?? "").Append(" nf:").Append(c.NumberFormatCode ?? "");
        }
        return sb.Length == 0 ? null : sb.ToString();
    }

    private static bool StringEquals(string? a, string? b)
        => string.Equals(a ?? "", b ?? "", StringComparison.Ordinal);

    private static void ParseAddress(string addr, out int row, out int col)
    {
        int i = 0;
        while (i < addr.Length && char.IsLetter(addr[i])) i++;
        var colStr = addr.Substring(0, i);
        var rowStr = addr.Substring(i);
        row = int.TryParse(rowStr, out var r) ? r : 0;
        col = ColumnNameToIndex(colStr);
    }

    private static int ColumnNameToIndex(string name)
    {
        int result = 0;
        foreach (var ch in name.ToUpperInvariant())
        {
            result = result * 26 + (ch - 'A' + 1);
        }
        return result;
    }

    private static string ColumnIndexToName(int index)
    {
        var sb = new StringBuilder();
        while (index > 0)
        {
            int rem = (index - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            index = (index - 1) / 26;
        }
        return sb.ToString();
    }
}