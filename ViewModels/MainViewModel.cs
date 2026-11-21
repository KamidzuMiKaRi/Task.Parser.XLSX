using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using Task.Parser.XLSX.Models;
using Task.Parser.XLSX.Services;
using Task.Parser.XLSX.Xls;

namespace Task.Parser.XLSX.ViewModels;

/// <summary>
/// Главная модель представления: загрузка .xls в память, изменение строк,
/// добавление, удаление, фильтрация по столбцам, сохранение отчёта (CSV).
/// Входной файл только читается — все изменения происходят в памяти.
/// </summary>
public sealed class MainViewModel : ViewModelBase
{
    /// <summary>Ошибка при выполнении действия; окно показывает её сообщение.</summary>
    public event EventHandler<string>? ErrorOccurred;

    /// <summary>Все строки таблицы (со всеми изменениями, сделанными в памяти).</summary>
    public ObservableCollection<TableRow> Rows { get; } = new();

    /// <summary>
    /// Представление для DataGrid (с фильтрацией). Пересоздаётся после загрузки
    /// файла и после изменения набора строк, поэтому грид всегда получает свежий
    /// источник данных и отображает содержимое.
    /// </summary>
    private ICollectionView? _view;

    public ICollectionView? View
    {
        get => _view;
        private set
        {
            _view = value;
            OnPropertyChanged();
        }
    }

    /// <summary>Фильтры по каждому столбцу (первая строка файла — заголовки).</summary>
    public ObservableCollection<ColumnFilterViewModel> Filters { get; } = new();

    private IReadOnlyList<string> _headers = Array.Empty<string>();

    /// <summary>Заголовки столбцов; при изменении уведомляет окно о пересоздании колонок.</summary>
    public IReadOnlyList<string> Headers
    {
        get => _headers;
        private set
        {
            _headers = value;
            OnPropertyChanged();
        }
    }

    private string _sheetName = string.Empty;

    public string SheetName
    {
        get => _sheetName;
        private set { _sheetName = value; OnPropertyChanged(); }
    }

    private string _sourceFileName = string.Empty;

    /// <summary>Имя загруженного файла (только для отображения).</summary>
    public string SourceFileName
    {
        get => _sourceFileName;
        private set { _sourceFileName = value; OnPropertyChanged(); }
    }

    private string _statusText = "Файл не загружен. Нажмите «Загрузить файл…»";

    public string StatusText
    {
        get => _statusText;
        private set { _statusText = value; OnPropertyChanged(); }
    }

    public ICommand LoadFileCommand { get; }
    public ICommand AddRowCommand { get; }
    public ICommand DeleteRowsCommand { get; }
    public ICommand SaveReportCommand { get; }
    public ICommand ClearFiltersCommand { get; }

    public MainViewModel()
    {
        LoadFileCommand = new RelayCommand(_ => LoadFile());
        AddRowCommand = new RelayCommand(_ => AddRow(), _ => Headers.Count > 0);
        DeleteRowsCommand = new RelayCommand(DeleteRows, _ => Headers.Count > 0 && Rows.Count > 0);
        SaveReportCommand = new RelayCommand(_ => SaveReport(), _ => Headers.Count > 0 && Rows.Count > 0);
        ClearFiltersCommand = new RelayCommand(_ => ClearFilters(), _ => Filters.Count > 0);
    }

    /// <summary>Показывает диалог выбора файла и загружает его в память.</summary>
    private void LoadFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Открыть файл Excel",
            Filter = "Книга Excel 97-2003 (*.xls)|*.xls|Все файлы (*.*)|*.*",
        };
        if (dialog.ShowDialog() != true)
            return;

        LoadFrom(dialog.FileName);
    }

    /// <summary>Загружает .xls файл в память (файл не изменяется).</summary>
    public void LoadFrom(string path)
    {
        try
        {
            XlsTable table;
            using (var document = Ole2CompoundFile.Open(path))
            {
                table = XlsWorkbookReader.ReadTable(document.GetStream("Workbook"));
            }

            Headers = table.Headers;
            SheetName = table.SheetName;

            Rows.Clear();
            foreach (var row in table.Rows)
            {
                var tableRow = new TableRow(table.Headers.Count);
                for (int c = 0; c < table.Headers.Count; c++)
                    tableRow[c] = row[c];
                Rows.Add(tableRow);
            }

            Filters.Clear();
            for (int c = 0; c < table.Headers.Count; c++)
            {
                var filter = new ColumnFilterViewModel(c, table.Headers[c]);
                filter.Changed += (_, _) => RefreshView();
                Filters.Add(filter);
            }

            SourceFileName = path;

            // Свежее представление по уже наполненной коллекции: DataGrid получит
            // новый ItemsSource и отобразит все строки.
            View = new ListCollectionView(Rows) { Filter = FilterRow };
            UpdateStatus();
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke(this, ex.Message);
        }
    }

    /// <summary>Добавляет новую (пустую) строку в конец таблицы.</summary>
    private void AddRow()
    {
        var row = new TableRow(Headers.Count);
        Rows.Add(row);
        EnsureView();
        RefreshView();
    }

    /// <summary>Удаляет выбранные строки (параметр — коллекция выделения DataGrid).</summary>
    private void DeleteRows(object? selected)
    {
        if (selected is not System.Collections.IList items || items.Count == 0)
        {
            StatusText = "Выберите строки для удаления.";
            return;
        }

        foreach (var item in items.Cast<TableRow>().ToList())
            Rows.Remove(item);
        EnsureView();
        RefreshView();
    }

    /// <summary>Сохраняет отчёт (CSV): заголовки и видимые строки со всеми изменениями.</summary>
    private void SaveReport()
    {
        var dialog = new SaveFileDialog
        {
            Title = "Сохранить отчёт (CSV)",
            Filter = "CSV (*.csv)|*.csv|Текстовый файл (*.txt)|*.txt",
            FileName = "отчёт.csv",
            DefaultExt = ".csv",
            AddExtension = true,
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            // Выгружаются строки, прошедшие фильтры (View), с внесёнными изменениями.
            var rowsToSave = View?.Cast<TableRow>() ?? Rows;
            CsvReportWriter.Save(dialog.FileName, Headers, rowsToSave);
            int count = rowsToSave.Count();
            StatusText = $"Отчёт сохранён: {dialog.FileName} (строк: {count})";
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke(this, ex.Message);
        }
    }

    private void ClearFilters()
    {
        foreach (var filter in Filters)
            filter.Query = string.Empty;
        RefreshView();
    }

    /// <summary>Гарантирует наличие представления (если оно ещё не создано).</summary>
    private void EnsureView()
    {
        if (View is null && Headers.Count > 0)
            View = new ListCollectionView(Rows) { Filter = FilterRow };
    }

    private void RefreshView()
    {
        EnsureView();
        View?.Refresh();
        UpdateStatus();
    }

    private bool FilterRow(object item)
    {
        if (item is not TableRow row)
            return false;

        foreach (var filter in Filters)
        {
            if (!filter.Matches(row[filter.ColumnIndex]))
                return false;
        }
        return true;
    }

    private void UpdateStatus()
    {
        int total = Rows.Count;
        int shown = View?.Cast<object>().Count() ?? 0;
        string file = string.IsNullOrEmpty(SourceFileName)
            ? "файл не загружен"
            : System.IO.Path.GetFileName(SourceFileName);
        StatusText = $"{file} | Лист «{SheetName}» | Строк: {total} | Показано: {shown}";
    }
}
