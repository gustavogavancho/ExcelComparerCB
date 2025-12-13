namespace ExcelComparerCB;

public enum DiffKind { Added, Removed, Modified }

public sealed record DiffItem(
    string Sheet,
    string Address,
    DiffKind Kind,
    string What,
    string? Before,
    string? After
);

public sealed class ComparisonOptions
{
    public bool CompareValues { get; set; } = true;
    public bool CompareFormulas { get; set; } = true;
    public bool IncludeHiddenSheets { get; set; } = true;

    // Workbook level
    public bool CompareSheetOrder { get; set; } = true;
    public bool CompareWorkbookProperties { get; set; } = false; // optional, not implemented fully

    // Worksheet level
    public bool CompareUsedRange { get; set; } = true; // approximate by max row/col with data
    public bool CompareValidations { get; set; } = true;
    public bool CompareConditionalFormats { get; set; } = false; // optional, may be heavy
    public bool CompareHiddenRowsCols { get; set; } = false; // optional

    // Cell level
    public bool CompareCellFormat { get; set; } = false; // NumberFormat / StyleId
}

public sealed record ProgressInfo(int Percent, string Message);

public sealed class ComparisonResult
{
    public List<DiffItem> Diffs { get; } = new();
}