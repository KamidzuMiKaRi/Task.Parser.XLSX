using System.Globalization;
using System.IO;
using System.Text;

namespace Task.Parser.XLSX.Xls;

/// <summary>
/// Парсер бинарного формата BIFF8 (Excel 97-2003) внутри потока Workbook
/// OLE2-документа. Поддерживает SST (общие строки), ячейки LABELSST, LABEL,
/// NUMBER, RK, MULRK, BLANK, MULBLANK и FORMULA.
///
/// Метод чтения листа:
/// 1) первая строка листа определяет число столбцов и заголовки столбцов;
/// 2) данные начинаются со второй строки;
/// 3) парсинг продолжается, пока строки не пустые (первая пустая строка
///    считается концом таблицы).
/// </summary>
public static class XlsWorkbookReader
{
    /// <summary>
    /// Кодировка ANSI-строк BIFF (для русской локали Excel — Windows-1251).
    /// В .NET Core+ кодовые страницы недоступны без регистрации провайдера.
    /// </summary>
    private static readonly Encoding AnsiEncoding;

    static XlsWorkbookReader()
    {
        // Сборка System.Text.Encoding.CodePages входит в общий фреймворк .NET.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        AnsiEncoding = Encoding.GetEncoding(1251);
    }

    private const ushort Bof = 0x0809;
    private const ushort Eof = 0x000A;
    private const ushort Boundsheet = 0x0085;
    private const ushort Sst = 0x00FC;
    private const ushort Continue = 0x003C;
    private const ushort Labelsst = 0x00FD;
    private const ushort Label = 0x0204;
    private const ushort Number = 0x0203;
    private const ushort Rk = 0x027E;
    private const ushort MulRk = 0x00BD;
    private const ushort Blank = 0x0201;
    private const ushort MulBlank = 0x00BE;
    private const ushort Formula = 0x0006;
    private const ushort StringRecord = 0x0207;

    private readonly record struct Record(ushort Opcode, byte[] Data, int Offset);

    /// <summary>Читает первый лист книги и возвращает его таблицу.</summary>
    public static XlsTable ReadTable(byte[] workbookStream)
    {
        var records = EnumerateRecords(workbookStream).ToList();

        // Проход 1: имена листов (BOUNDSHEET) и таблица общих строк (SST + CONTINUE).
        var sheets = new List<(string Name, uint Offset)>();
        List<string> sharedStrings = new();

        for (int i = 0; i < records.Count; i++)
        {
            var rec = records[i];
            if (rec.Opcode == Boundsheet)
            {
                sheets.Add((ReadShortUnicodeString(rec.Data, 6), BitConverter.ToUInt32(rec.Data, 0)));
            }
            else if (rec.Opcode == Sst)
            {
                sharedStrings = ReadSharedStrings(records, i);
            }
        }

        if (sheets.Count == 0)
            throw new InvalidDataException("В книге не найдено ни одного листа.");

        // Проход 2: ячейки первого листа (начиная с его смещения в потоке).
        var sheet = sheets[0];
        int startIndex = -1;
        for (int i = 0; i < records.Count; i++)
        {
            if (records[i].Offset >= sheet.Offset)
            {
                startIndex = i;
                break;
            }
        }
        if (startIndex < 0)
            throw new InvalidDataException("Ошибка позиционирования на лист книги.");

        var cells = new SortedDictionary<int, SortedDictionary<int, string>>();

        for (int i = startIndex; i < records.Count; i++)
        {
            var rec = records[i];
            switch (rec.Opcode)
            {
                case Labelsst:
                    {
                        int row = BitConverter.ToUInt16(rec.Data, 0);
                        int col = BitConverter.ToUInt16(rec.Data, 2);
                        int isst = BitConverter.ToInt32(rec.Data, 6);
                        if (isst >= 0 && isst < sharedStrings.Count)
                            SetCell(cells, row, col, sharedStrings[isst]);
                        break;
                    }
                case Label:
                    {
                        int row = BitConverter.ToUInt16(rec.Data, 0);
                        int col = BitConverter.ToUInt16(rec.Data, 2);
                        SetCell(cells, row, col, ReadUnicodeString(rec.Data, 6));
                        break;
                    }
                case Number:
                    {
                        int row = BitConverter.ToUInt16(rec.Data, 0);
                        int col = BitConverter.ToUInt16(rec.Data, 2);
                        SetCell(cells, row, col, FormatNumber(BitConverter.ToDouble(rec.Data, 6)));
                        break;
                    }
                case Rk:
                    {
                        int row = BitConverter.ToUInt16(rec.Data, 0);
                        int col = BitConverter.ToUInt16(rec.Data, 2);
                        SetCell(cells, row, col, FormatNumber(DecodeRk(BitConverter.ToInt32(rec.Data, 6))));
                        break;
                    }
                case MulRk:
                    {
                        int row = BitConverter.ToUInt16(rec.Data, 0);
                        int colFirst = BitConverter.ToUInt16(rec.Data, 2);
                        int colLast = BitConverter.ToUInt16(rec.Data, rec.Data.Length - 2);
                        for (int col = colFirst; col <= colLast; col++)
                        {
                            // На ячейку приходится ixfe(2) + rk(4).
                            int idx = 4 + (col - colFirst) * 6 + 2;
                            SetCell(cells, row, col, FormatNumber(DecodeRk(BitConverter.ToInt32(rec.Data, idx))));
                        }
                        break;
                    }
                case Blank:
                    break;
                case Formula:
                    {
                        int row = BitConverter.ToUInt16(rec.Data, 0);
                        int col = BitConverter.ToUInt16(rec.Data, 2);
                        ulong raw = BitConverter.ToUInt64(rec.Data, 6);
                        if ((raw >> 32) == 0xFFFFFFFF)
                        {
                            // Нечисловой результат: старший байт кодирует тип
                            // (0 — строка, 1 — логическое, 2 — ошибка, 3 — пустая строка).
                            byte type = rec.Data[13];
                            if (type == 0 && i + 1 < records.Count && records[i + 1].Opcode == StringRecord)
                                SetCell(cells, row, col, ReadUnicodeString(records[i + 1].Data, 0));
                            else if (type == 1)
                                SetCell(cells, row, col, rec.Data[12] != 0 ? "ИСТИНА" : "ЛОЖЬ");
                        }
                        else
                        {
                            SetCell(cells, row, col, FormatNumber(BitConverter.ToDouble(rec.Data, 6)));
                        }
                        break;
                    }
                case Eof:
                    goto sheetDone;
            }
        }

        sheetDone:

        if (cells.Count == 0)
            throw new InvalidDataException("Лист не содержит данных.");

        // Метод парсинга: первая строка определяет число столбцов и заголовки.
        if (!cells.TryGetValue(0, out var headerCells) || headerCells.Count == 0)
            throw new InvalidDataException(
                "Лист не содержит заголовков: первая строка пуста.");

        int colCount = headerCells.Keys.Max() + 1;
        var headers = new string[colCount];
        for (int c = 0; c < colCount; c++)
            headers[c] = GetCell(headerCells, c);

        // Данные начинаются со второй строки; парсинг продолжается,
        // пока строки не пустые (первая пустая строка — конец таблицы).
        var rows = new List<IReadOnlyList<string>>();
        for (int row = 1; cells.TryGetValue(row, out var rowCells); row++)
        {
            var rowValues = new string[colCount];
            for (int c = 0; c < colCount; c++)
                rowValues[c] = GetCell(rowCells, c);
            rows.Add(rowValues);
        }

        return new XlsTable
        {
            SheetName = sheet.Name,
            Headers = headers,
            Rows = rows,
        };
    }

    private static IEnumerable<Record> EnumerateRecords(byte[] stream)
    {
        int pos = 0;
        while (pos + 4 <= stream.Length)
        {
            ushort opcode = BitConverter.ToUInt16(stream, pos);
            ushort length = BitConverter.ToUInt16(stream, pos + 2);
            if (pos + 4 + length > stream.Length)
                throw new InvalidDataException("BIFF: запись выходит за границы потока.");
            yield return new Record(opcode, stream[(pos + 4)..(pos + 4 + length)], pos);
            pos += 4 + length;
        }
    }

    private static void SetCell(
        SortedDictionary<int, SortedDictionary<int, string>> cells, int row, int col, string value)
    {
        if (!cells.TryGetValue(row, out var rowCells))
        {
            rowCells = new SortedDictionary<int, string>();
            cells[row] = rowCells;
        }
        rowCells[col] = value;
    }

    private static string GetCell(SortedDictionary<int, string> rowCells, int col)
    {
        return rowCells.TryGetValue(col, out var value) ? value : string.Empty;
    }

    /// <summary>
    /// Читает таблицу общих строк SST вместе с идущими следом записями CONTINUE.
    /// Если символьные данные строки пересекают границу записи, в начале следующей
    /// записи лежит байт-флаг ширины символов.
    /// </summary>
    private static List<string> ReadSharedStrings(List<Record> records, int sstIndex)
    {
        var parts = new List<byte[]> { records[sstIndex].Data };
        for (int i = sstIndex + 1; i < records.Count && records[i].Opcode == Continue; i++)
            parts.Add(records[i].Data);

        var stream = new SstStream(parts);
        if (!stream.TryReadInt32(out int total) || !stream.TryReadInt32(out int unique))
            return new List<string>();

        var strings = new List<string>(Math.Max(0, unique));
        for (int i = 0; i < unique; i++)
        {
            if (!stream.TryReadUInt16(out ushort cch) || !stream.TryReadByte(out byte flags))
                break;

            bool fHighByte = (flags & 0x01) != 0;
            bool fExtSt = (flags & 0x04) != 0;
            bool fRichSt = (flags & 0x08) != 0;

            if (fRichSt)
            {
                if (!stream.TryReadUInt16(out ushort cRun))
                    break;
                for (int r = 0; r < cRun && stream.TrySkip(4); r++) { }
            }

            if (fExtSt)
            {
                if (!stream.TryReadInt32(out int cbExtRst) || !stream.TrySkip(cbExtRst))
                    break;
            }

            strings.Add(stream.ReadString(cch, fHighByte));
        }

        return strings;
    }

    /// <summary>Читает строку ShortXLUnicodeString (однобайтовая длина; BOUNDSHEET).</summary>
    private static string ReadShortUnicodeString(byte[] data, int offset)
    {
        byte cch = data[offset];
        byte flags = data[offset + 1];
        bool fHighByte = (flags & 0x01) != 0;
        int start = offset + 2;
        if (fHighByte)
            return Encoding.Unicode.GetString(data, start, cch * 2);
        return AnsiEncoding.GetString(data, start, cch);
    }

    /// <summary>Читает строку XLUnicodeString (двухбайтовая длина; LABEL, STRING).</summary>
    private static string ReadUnicodeString(byte[] data, int offset)
    {
        ushort cch = BitConverter.ToUInt16(data, offset);
        byte flags = data[offset + 2];
        bool fHighByte = (flags & 0x01) != 0;
        int start = offset + 3;
        if (fHighByte)
            return Encoding.Unicode.GetString(data, start, cch * 2);
        return AnsiEncoding.GetString(data, start, cch);
    }

    /// <summary>
    /// Распаковка числа RK: если установлен бит 1 — 30-битное знаковое целое;
    /// иначе — старшие 30 бит 64-битного IEEE 754. Бит 0 означает деление на 100.
    /// </summary>
    private static double DecodeRk(int rk)
    {
        uint bits = unchecked((uint)rk);
        double value;
        if ((bits & 0x02) != 0)
        {
            // 30-битное знаковое целое: знак сохраняется арифметическим сдвигом.
            value = ((int)(bits & 0xFFFFFFFC)) >> 2;
        }
        else
        {
            // Старшие 30 значащих бит double; младшие 2 байта равны нулю.
            value = BitConverter.Int64BitsToDouble((long)(bits & 0xFFFFFFFC) << 32);
        }
        return (bits & 0x01) != 0 ? value / 100.0 : value;
    }

    /// <summary>Форматирует число для отображения и отчёта (целые — без дробной части).</summary>
    private static string FormatNumber(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
            return string.Empty;
        if (Math.Abs(value) < 1e15 && value == Math.Truncate(value))
            return ((long)value).ToString(CultureInfo.CurrentCulture);
        return value.ToString(CultureInfo.CurrentCulture);
    }

    /// <summary>Поток байт секции SST, прозрачно пересекающий границы записей CONTINUE.</summary>
    private sealed class SstStream
    {
        private readonly byte[][] _parts;
        private int _part;
        private int _pos;

        public SstStream(IReadOnlyList<byte[]> parts)
        {
            _parts = parts.ToArray();
        }

        private bool MoveNextPart()
        {
            if (_part + 1 >= _parts.Length)
                return false;
            _part++;
            _pos = 0;
            return true;
        }

        public bool TryReadByte(out byte value)
        {
            if (_pos >= _parts[_part].Length && !MoveNextPart())
            {
                value = 0;
                return false;
            }
            value = _parts[_part][_pos++];
            return true;
        }

        public bool TryReadUInt16(out ushort value)
        {
            if (!TryReadByte(out byte lo) || !TryReadByte(out byte hi))
            {
                value = 0;
                return false;
            }
            value = (ushort)(lo | (hi << 8));
            return true;
        }

        public bool TryReadInt32(out int value)
        {
            if (!TryReadUInt16(out ushort lo) || !TryReadUInt16(out ushort hi))
            {
                value = 0;
                return false;
            }
            value = lo | (hi << 16);
            return true;
        }

        public bool TrySkip(int count)
        {
            while (count > 0)
            {
                if (_pos >= _parts[_part].Length && !MoveNextPart())
                    return false;
                int take = Math.Min(count, _parts[_part].Length - _pos);
                _pos += take;
                count -= take;
            }
            return true;
        }

        /// <summary>
        /// Читает строку из cch символов. Если символьные данные пересекают границу
        /// записи, в начале следующей части лежит байт-флаг ширины символов (fHighByte).
        /// </summary>
        public string ReadString(int cch, bool fHighByte)
        {
            var sb = new StringBuilder(cch);
            int remaining = cch;
            while (remaining > 0)
            {
                if (_pos >= _parts[_part].Length)
                {
                    if (!MoveNextPart())
                        break;
                    // Флаг ширины символов продолжения — первый байт новой записи.
                    fHighByte = (_parts[_part][_pos++] & 0x01) != 0;
                    continue;
                }
                if (fHighByte)
                {
                    if (_pos + 2 > _parts[_part].Length)
                    {
                        // Символ разорван границей: старший байт кода лежит уже
                        // в следующей части (это данные, а не флаг).
                        byte lo = _parts[_part][_pos++];
                        if (!MoveNextPart())
                            break;
                        byte hi = _parts[_part][_pos++];
                        sb.Append((char)(lo | (hi << 8)));
                    }
                    else
                    {
                        byte lo = _parts[_part][_pos++];
                        byte hi = _parts[_part][_pos++];
                        sb.Append((char)(lo | (hi << 8)));
                    }
                }
                else
                {
                    sb.Append(AnsiEncoding.GetString(_parts[_part], _pos++, 1));
                }
                remaining--;
            }
            return sb.ToString();
        }
    }
}
