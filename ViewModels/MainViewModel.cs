using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using Task.Parser.XLSX.Models;
using Task.Parser.XLSX.Xls;

namespace Task.Parser.XLSX.ViewModels;

/// <summary>
/// Главная модель представления: загрузка .xls в память, изменение строк,
/// фильтрация по столбцам. Входной файл только читается — все изменения
/// происходят в памяти.
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
    public ICommand ClearFiltersCommand { get; }

    public MainViewModel()
    {
        LoadFileCommand = new RelayCommand(_ => LoadFile());
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
