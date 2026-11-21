using System.IO;
using System.Text;

namespace Task.Parser.XLSX.Services;

/// <summary>
/// Сохраняет отчёт: таблицу (заголовки + строки, все изменения уже применены
/// в памяти) в текстовый файл CSV. Разделитель — точка с запятой (русская
/// локаль Excel), кодировка UTF-8 с BOM.
/// </summary>
public static class CsvReportWriter
{
    public static void Save(string path, IReadOnlyList<string> headers, IEnumerable<IReadOnlyList<string>> rows)
    {
        using var writer = new StreamWriter(
            path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        writer.WriteLine(string.Join(";", headers.Select(Escape)));

        foreach (var row in rows)
            writer.WriteLine(string.Join(";", row.Select(Escape)));
    }

    private static string Escape(string value)
    {
        bool needQuotes = value.Contains(';') || value.Contains('"') ||
                          value.Contains('\n') || value.Contains('\r');
        if (!needQuotes)
            return value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}
