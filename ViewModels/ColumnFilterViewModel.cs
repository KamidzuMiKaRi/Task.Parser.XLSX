using System.Collections.ObjectModel;

namespace Task.Parser.XLSX.ViewModels;

/// <summary>Операторы фильтра для одного столбца.</summary>
public static class FilterOperators
{
    public const string Contains = "Содержит";
    public const string Equal = "Равно";
}

/// <summary>
/// Фильтр по одному столбцу: оператор (Содержит / Равно) и значение.
/// Любое изменение уведомляет подписчиков (событие Changed).
/// </summary>
public sealed class ColumnFilterViewModel : ViewModelBase
{
    public event EventHandler? Changed;

    public int ColumnIndex { get; }

    public string Header { get; }

    /// <summary>Доступные операторы (в порядке отображения).</summary>
    public ObservableCollection<string> Operators { get; } = new()
    {
        FilterOperators.Contains,
        FilterOperators.Equal,
    };

    private string _selectedOperator = FilterOperators.Contains;

    public string SelectedOperator
    {
        get => _selectedOperator;
        set
        {
            if (_selectedOperator == value)
                return;
            _selectedOperator = value;
            OnPropertyChanged();
            Changed?.Invoke(this, EventArgs.Empty);
        }
    }

    private string _query = string.Empty;

    public string Query
    {
        get => _query;
        set
        {
            if (string.Equals(_query, value, StringComparison.Ordinal))
                return;
            _query = value;
            OnPropertyChanged();
            Changed?.Invoke(this, EventArgs.Empty);
        }
    }

    public ColumnFilterViewModel(int columnIndex, string header)
    {
        ColumnIndex = columnIndex;
        Header = header;
    }

    /// <summary>Применяет условие к значению ячейки; true — строка проходит фильтр.</summary>
    public bool Matches(string cellValue)
    {
        string query = _query.Trim();
        if (query.Length == 0)
            return true;

        if (string.Equals(_selectedOperator, FilterOperators.Equal, StringComparison.Ordinal))
            return string.Equals(cellValue, query, StringComparison.OrdinalIgnoreCase);

        return cellValue.Contains(query, StringComparison.OrdinalIgnoreCase);
    }
}
