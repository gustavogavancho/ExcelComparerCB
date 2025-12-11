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
}

public sealed record ProgressInfo(int Percent, string Message);

public sealed class ComparisonResult
{
    public List<DiffItem> Diffs { get; } = new();
}