namespace Task.Parser.XLSX.Xls;

/// <summary>
/// Таблица, прочитанная из листа книги Excel: заголовки и строки данных
/// (все значения — строки, числа отформатированы для отображения).
/// </summary>
public sealed class XlsTable
{
    /// <summary>Имя листа (например, "Клиенты").</summary>
    public string SheetName { get; init; } = string.Empty;

    /// <summary>Заголовки столбцов (первая строка листа).</summary>
    public IReadOnlyList<string> Headers { get; init; } = Array.Empty<string>();

    /// <summary>Строки данных (без строки заголовков).</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; init; } = Array.Empty<IReadOnlyList<string>>();
}
