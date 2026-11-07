using System.IO;
using System.Text;

namespace Task.Parser.XLSX.Xls;

/// <summary>
/// Минимальный ридер контейнера OLE2 (Compound File Binary), в котором хранится
/// поток Workbook файла .xls (Excel 97-2003). Входной файл читается целиком
/// в память и только для чтения.
/// </summary>
public sealed class Ole2CompoundFile : IDisposable
{
    private const uint FreeSector = 0xFFFFFFFF;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSector = 0xFFFFFFFD;
    private const uint DifatSector = 0xFFFFFFFC;
    private const uint MiniFatSector = 0xFFFFFFFB;

    private static readonly byte[] Signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    private readonly byte[] _file;
    private readonly int _sectorSize;
    private readonly int _miniSectorSize;
    private readonly uint _miniStreamCutoff;

    private uint[] _fat = Array.Empty<uint>();
    private uint[] _miniFat = Array.Empty<uint>();
    private byte[] _miniStream = Array.Empty<byte>();

    private readonly List<(string Name, int Type, uint StartSector, long Size)> _entries = new();

    private Ole2CompoundFile(byte[] file)
    {
        _file = file;
        (_sectorSize, _miniSectorSize, _miniStreamCutoff) = ReadHeader();
        ParseDirectory();
        ReadMiniStream();
    }

    /// <summary>Открывает OLE2-документ (только чтение).</summary>
    public static Ole2CompoundFile Open(string path)
    {
        var bytes = File.ReadAllBytes(path);
        if (!Signature.SequenceEqual(bytes.Take(8)))
            throw new InvalidDataException("Файл не является документом OLE2 (.xls).");
        return new Ole2CompoundFile(bytes);
    }

    public void Dispose()
    {
        // Содержимое держится в памяти; освобождать нечего.
    }

    private (int Sector, int Mini, uint Cutoff) ReadHeader()
    {
        if (_file.Length < 512)
            throw new InvalidDataException("Файл повреждён: слишком мал для OLE2-документа.");

        ushort version = BitConverter.ToUInt16(_file, 26);
        ushort byteOrder = BitConverter.ToUInt16(_file, 28);
        ushort sectorShift = BitConverter.ToUInt16(_file, 30);
        ushort miniShift = BitConverter.ToUInt16(_file, 32);
        if (byteOrder != 0xFFFE)
            throw new InvalidDataException("OLE2: порядок байт не поддерживается.");

        int sectorSize = 1 << sectorShift;
        int miniSize = 1 << miniShift;
        if (version != 3 && version != 4)
            throw new InvalidDataException($"OLE2: неподдерживаемая версия документа ({version}).");

        uint numFatSectors = BitConverter.ToUInt32(_file, 44);
        uint firstDirSector = BitConverter.ToUInt32(_file, 48);
        uint miniCutoff = BitConverter.ToUInt32(_file, 56);
        uint firstMiniFatSector = BitConverter.ToUInt32(_file, 60);
        uint numMiniFatSectors = BitConverter.ToUInt32(_file, 64);
        uint firstDifatSector = BitConverter.ToUInt32(_file, 68);
        uint numDifatSectors = BitConverter.ToUInt32(_file, 72);

        var difat = new List<uint>(109);
        for (int i = 0; i < 109; i++)
            difat.Add(BitConverter.ToUInt32(_file, 76 + i * 4));

        // DIFAT может продолжаться отдельными секторами; последний элемент
        // каждого такого сектора указывает на следующий сектор цепочки.
        uint current = firstDifatSector;
        uint read = 0;
        while (current < 0xFFFFFFF0 && read < numDifatSectors)
        {
            long offset = SectorOffset(current);
            int perSector = sectorSize / 4;
            for (int i = 0; i < perSector - 1; i++)
                difat.Add(BitConverter.ToUInt32(_file, (int)offset + i * 4));
            current = BitConverter.ToUInt32(_file, (int)offset + (perSector - 1) * 4);
            read++;
        }

        // Собираем FAT из секторов, перечисленных в DIFAT.
        int expectedFatEntries = checked((int)numFatSectors * (sectorSize / 4));
        var fat = new List<uint>(expectedFatEntries);
        foreach (uint sector in difat)
        {
            if (sector == FreeSector || sector == EndOfChain || sector == FatSector ||
                sector == DifatSector || sector == MiniFatSector)
                continue;
            AppendFatSector(fat, sector, sectorSize);
        }

        if (fat.Count < expectedFatEntries)
            throw new InvalidDataException(
                $"OLE2: не удалось прочитать все секторы FAT (получено {fat.Count} из {expectedFatEntries}).");

        _fat = fat.ToArray();
        return (sectorSize, miniSize, miniCutoff);
    }

    private void AppendFatSector(List<uint> fat, uint sector, int sectorSize)
    {
        // Внимание: во время чтения заголовка поле _sectorSize ещё не присвоено,
        // поэтому смещение считается от локального размера сектора.
        long offset = 512L + sector * sectorSize;
        int perSector = sectorSize / 4;
        for (int i = 0; i < perSector; i++)
            fat.Add(BitConverter.ToUInt32(_file, (int)offset + i * 4));
    }

    private long SectorOffset(uint sector)
    {
        // Сектор 0 расположен сразу после 512-байтного заголовка.
        return 512L + sector * _sectorSize;
    }

    private IReadOnlyList<uint> Chain(uint start)
    {
        var chain = new List<uint>();
        uint s = start;
        int guard = 0;
        while (s < 0xFFFFFFF0 && s < _fat.Length && guard++ < 1_000_000)
        {
            chain.Add(s);
            s = _fat[s];
        }
        return chain;
    }

    private byte[] ReadChain(uint start, long size, string streamName)
    {
        var chain = Chain(start);

        // Проверка целостности: цепочка секторов должна покрывать объявленный
        // размер потока (иначе читатель вернул бы «тихие» нули).
        long expectedSectors = (size + _sectorSize - 1) / _sectorSize;
        if (chain.Count < expectedSectors)
            throw new InvalidDataException(
                $"OLE2: поток \"{streamName}\": цепочка секторов короче размера потока " +
                $"({chain.Count} < {expectedSectors}).");

        var result = new byte[size];
        long written = 0;
        foreach (uint sector in chain)
        {
            if (written >= size)
                break;
            int take = (int)Math.Min(_sectorSize, size - written);
            Array.Copy(_file, SectorOffset(sector), result, written, take);
            written += take;
        }
        return result;
    }

    private void ParseDirectory()
    {
        uint firstDir = BitConverter.ToUInt32(_file, 48);
        var chain = Chain(firstDir);

        // Размер каталога нигде не объявлен: читаем цепочку до её конца.
        var dirBytes = new byte[checked(chain.Count * _sectorSize)];
        long written = 0;
        foreach (uint sector in chain)
        {
            Array.Copy(_file, SectorOffset(sector), dirBytes, written, _sectorSize);
            written += _sectorSize;
        }

        int pos = 0;
        while (pos + 128 <= dirBytes.Length)
        {
            ushort nameLength = BitConverter.ToUInt16(dirBytes, pos + 64);
            byte type = dirBytes[pos + 66];
            if (nameLength == 0 && type == 0)
                break;

            string name = nameLength > 2
                ? Encoding.Unicode.GetString(dirBytes, pos, nameLength - 2)
                : string.Empty;
            uint start = BitConverter.ToUInt32(dirBytes, pos + 116);
            long size = BitConverter.ToInt64(dirBytes, pos + 120);
            _entries.Add((name, type, start, size));
            pos += 128;
        }
    }

    private void ReadMiniStream()
    {
        uint numMiniFat = BitConverter.ToUInt32(_file, 64);
        if (numMiniFat == 0)
            return;

        uint firstMiniFat = BitConverter.ToUInt32(_file, 60);
        var miniFatBytes = ReadChain(firstMiniFat, numMiniFat * (uint)_sectorSize, "мини-FAT");
        _miniFat = new uint[miniFatBytes.Length / 4];
        for (int i = 0; i < _miniFat.Length; i++)
            _miniFat[i] = BitConverter.ToUInt32(miniFatBytes, i * 4);

        var root = _entries.FirstOrDefault(e => e.Type == 5);
        if (root == default)
            return;
        _miniStream = ReadChain(root.StartSector, root.Size, "мини-поток");
    }

    /// <summary>Возвращает содержимое именованного потока (например, "Workbook").</summary>
    public byte[] GetStream(string name)
    {
        var entry = _entries.FirstOrDefault(e =>
            string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase));
        if (entry == default)
            throw new InvalidDataException($"OLE2: поток \"{name}\" не найден. " +
                $"Доступные потоки: {string.Join(", ", _entries.Select(e => e.Name))}");

        var result = new byte[entry.Size];
        if (entry.Size < _miniStreamCutoff && _miniFat.Length > 0)
        {
            // Поток лежит в мини-потоке (секторы внутри Root Entry).
            uint s = entry.StartSector;
            long written = 0;
            int guard = 0;
            while (s < 0xFFFFFFF0 && s < _miniFat.Length && guard++ < 1_000_000)
            {
                int take = (int)Math.Min(_miniSectorSize, entry.Size - written);
                Array.Copy(_miniStream, s * _miniSectorSize, result, written, take);
                written += take;
                if (written >= entry.Size)
                    break;
                s = _miniFat[s];
            }
            if (written < entry.Size)
                throw new InvalidDataException(
                    $"OLE2: поток \"{name}\": мини-цепочка секторов короче размера потока.");
        }
        else
        {
            Array.Copy(ReadChain(entry.StartSector, entry.Size, name), result, entry.Size);
        }
        return result;
    }
}
