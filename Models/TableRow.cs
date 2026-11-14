using System.Collections;
using System.ComponentModel;

namespace Task.Parser.XLSX.Models;

/// <summary>
/// Строка таблицы в памяти приложения. Значения ячеек доступны по индексу;
/// изменение ячейки уведомляет интерфейс (Binding Path="[i]").
/// </summary>
public sealed class TableRow : INotifyPropertyChanged, IReadOnlyList<string>
{
    private readonly string[] _cells;

    public TableRow(int columnCount)
    {
        _cells = new string[columnCount];
    }

    /// <summary>Количество ячеек в строке.</summary>
    public int Count => _cells.Length;

    public string this[int index]
    {
        get => _cells[index];
        set
        {
            if (string.Equals(_cells[index], value, StringComparison.Ordinal))
                return;
            _cells[index] = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs($"Item[{index}]"));
        }
    }

    public IEnumerator<string> GetEnumerator() => ((IEnumerable<string>)_cells).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => _cells.GetEnumerator();

    public event PropertyChangedEventHandler? PropertyChanged;
}
