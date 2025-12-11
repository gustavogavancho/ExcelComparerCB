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
        // CPU/I/O bound: lo llevamos a background thread
        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();

            var result = new ComparisonResult();
            progress?.Report(new ProgressInfo(1, "Leyendo estructura de libros..."));

            using var docA = SpreadsheetDocument.Open(fileA, false);
            using var docB = SpreadsheetDocument.Open(fileB, false);

            var wbA = ReadWorkbook(docA);
            var wbB = ReadWorkbook(docB);

            // 1) Diff de hojas
            progress?.Report(new ProgressInfo(5, "Comparando hojas..."));
            DiffSheets(wbA, wbB, result, options);

            // 2) Diff de celdas por hoja
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

                progress?.Report(new ProgressInfo(pct, $"Comparando celdas: {sheetName} ({i + 1}/{total})"));

                // Si una hoja existe solo en A o solo en B, igualmente reportamos (ya se reportó como sheet add/remove)
                // Para MVP, solo comparamos celdas si existe en ambos.
                if (!wbA.SheetsByName.TryGetValue(sheetName, out var sA) ||
                    !wbB.SheetsByName.TryGetValue(sheetName, out var sB))
                {
                    continue;
                }

                // Excluir ocultas si el usuario lo pide
                if (!options.IncludeHiddenSheets && (sA.Hidden || sB.Hidden))
                    continue;

                var cellsA = ReadCells(docA, sA);
                var cellsB = ReadCells(docB, sB);

                DiffCells(sheetName, cellsA, cellsB, result, options);
            }

            progress?.Report(new ProgressInfo(100, "Finalizado."));
            return result;

        }, ct);
    }

    // -------- Workbook reading --------

    private sealed class WorkbookInfo
    {
        public Dictionary<string, SheetInfo> SheetsByName { get; } = new(StringComparer.OrdinalIgnoreCase);
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

    // -------- Cell reading + diff --------

    private sealed class CellInfo
    {
        public string? ValueText { get; init; }     // normalizado a texto
        public string? FormulaText { get; init; }   // fórmula tal cual
    }

    private static Dictionary<string, CellInfo> ReadCells(SpreadsheetDocument doc, SheetInfo sheet)
    {
        var wbPart = doc.WorkbookPart!;
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.RelId);

        var sharedStrings = wbPart.SharedStringTablePart?.SharedStringTable;
        var sheetData = wsPart.Worksheet.Elements<SheetData>().FirstOrDefault();

        var dict = new Dictionary<string, CellInfo>(StringComparer.OrdinalIgnoreCase);
        if (sheetData is null) return dict;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var addr = cell.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(addr)) continue;

                var formula = cell.CellFormula?.Text;
                var val = ReadCellValueAsText(cell, sharedStrings);

                // Guardamos solo celdas relevantes (tocadas)
                if (val is null && formula is null) continue;

                dict[addr] = new CellInfo
                {
                    ValueText = val,
                    FormulaText = formula
                };
            }
        }

        return dict;
    }

    private static string? ReadCellValueAsText(Cell cell, SharedStringTable? sst)
    {
        // Nota: este "ValueText" puede ser:
        // - valor literal (para constantes)
        // - valor cacheado en CellValue (para fórmulas, si existe)
        // No intentamos recalcular.

        if (cell.CellValue is null) return null;
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

    private static void DiffCells(
        string sheetName,
        Dictionary<string, CellInfo> a,
        Dictionary<string, CellInfo> b,
        ComparisonResult result,
        ComparisonOptions options)
    {
        var keys = a.Keys.Union(b.Keys, StringComparer.OrdinalIgnoreCase);

        foreach (var addr in keys.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            var hasA = a.TryGetValue(addr, out var cA);
            var hasB = b.TryGetValue(addr, out var cB);

            if (hasA && !hasB)
            {
                // Removed
                result.Diffs.Add(new DiffItem(sheetName, addr, DiffKind.Removed, "Cell", Summarize(cA, options), null));
                continue;
            }
            if (!hasA && hasB)
            {
                // Added
                result.Diffs.Add(new DiffItem(sheetName, addr, DiffKind.Added, "Cell", null, Summarize(cB, options)));
                continue;
            }

            // ambos
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

            foreach (var ch in changes)
                result.Diffs.Add(ch);
        }
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
        return sb.Length == 0 ? null : sb.ToString();
    }

    private static bool StringEquals(string? a, string? b)
        => string.Equals(a ?? "", b ?? "", StringComparison.Ordinal);
}