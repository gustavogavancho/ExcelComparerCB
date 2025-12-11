using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExcelComparerCB;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<DiffItem> _diffs = new();
    private readonly ICollectionView _diffsView;
    private CancellationTokenSource? _cts;

    public MainWindow()
    {
        InitializeComponent();

        _diffsView = CollectionViewSource.GetDefaultView(_diffs);
        _diffsView.Filter = FilterDiffs;

        GridDiffs.ItemsSource = _diffsView;
        Prog.Value = 0;
    }

    private void BtnBrowseA_Click(object sender, RoutedEventArgs e)
        => TxtFileA.Text = PickExcelFile();

    private void BtnBrowseB_Click(object sender, RoutedEventArgs e)
        => TxtFileB.Text = PickExcelFile();

    private static string PickExcelFile()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
            CheckFileExists = true
        };
        return dlg.ShowDialog() == true ? dlg.FileName : "";
    }

    private async void BtnCompare_Click(object sender, RoutedEventArgs e)
    {
        var fileA = TxtFileA.Text?.Trim();
        var fileB = TxtFileB.Text?.Trim();

        if (string.IsNullOrWhiteSpace(fileA) || string.IsNullOrWhiteSpace(fileB))
        {
            MessageBox.Show("Selecciona ambos archivos.", "Excel Diff", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        BtnCompare.IsEnabled = false;
        BtnCancel.IsEnabled = true;
        TxtStatus.Text = "Comparando...";
        Prog.Value = 0;
        _diffs.Clear();
        TreeSummary.Items.Clear();

        _cts = new CancellationTokenSource();

        var options = new ComparisonOptions
        {
            CompareValues = ChkCompareValues.IsChecked == true,
            CompareFormulas = ChkCompareFormulas.IsChecked == true,
            IncludeHiddenSheets = ChkIncludeHiddenSheets.IsChecked == true
        };

        var progress = new Progress<ProgressInfo>(p =>
        {
            Prog.Value = p.Percent;
            TxtStatus.Text = p.Message;
        });

        try
        {
            var comparer = new ExcelComparer();
            var result = await comparer.CompareAsync(fileA, fileB, options, progress, _cts.Token);

            // Fill grid
            foreach (var d in result.Diffs)
                _diffs.Add(d);

            // Fill tree summary
            TreeSummary.Items.Add(BuildSummaryTree(result));

            TxtStatus.Text = $"Listo. Cambios: {_diffs.Count}";
            Prog.Value = 100;
        }
        catch (OperationCanceledException)
        {
            TxtStatus.Text = "Cancelado.";
        }
        catch (Exception ex)
        {
            TxtStatus.Text = "Error.";
            MessageBox.Show(ex.Message, "Excel Diff - Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnCompare.IsEnabled = true;
            BtnCancel.IsEnabled = false;
            _cts?.Dispose();
            _cts = null;
        }
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
        => _cts?.Cancel();

    private void TxtFilter_TextChanged(object sender, TextChangedEventArgs e)
        => _diffsView.Refresh();

    private bool FilterDiffs(object obj)
    {
        if (obj is not DiffItem d) return false;

        var q = TxtFilter.Text?.Trim();
        if (string.IsNullOrWhiteSpace(q)) return true;

        q = q.ToLowerInvariant();

        return (d.Sheet?.ToLowerInvariant().Contains(q) ?? false)
            || (d.Address?.ToLowerInvariant().Contains(q) ?? false)
            || d.Kind.ToString().ToLowerInvariant().Contains(q)
            || (d.What?.ToLowerInvariant().Contains(q) ?? false)
            || (d.Before?.ToLowerInvariant().Contains(q) ?? false)
            || (d.After?.ToLowerInvariant().Contains(q) ?? false);
    }

    private static TreeViewItem BuildSummaryTree(ComparisonResult result)
    {
        var root = new TreeViewItem { Header = $"Workbook diff (total: {result.Diffs.Count})", IsExpanded = true };

        var bySheet = result.Diffs
            .GroupBy(d => d.Sheet)
            .OrderBy(g => g.Key);

        foreach (var g in bySheet)
        {
            var added = g.Count(x => x.Kind == DiffKind.Added);
            var removed = g.Count(x => x.Kind == DiffKind.Removed);
            var modified = g.Count(x => x.Kind == DiffKind.Modified);

            var sheetNode = new TreeViewItem
            {
                Header = $"{g.Key}   (+{added}  ~{modified}  -{removed})",
                IsExpanded = false
            };
            root.Items.Add(sheetNode);
        }

        return root;
    }
}