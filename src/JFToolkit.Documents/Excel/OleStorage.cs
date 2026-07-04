using System.Runtime.InteropServices;

namespace JFToolkit.Documents.Excel;

/// <summary>
/// Minimal OLE2 Compound File Binary (CFB) reader/writer.
/// 
/// Supports v3 (512-byte sectors) — the only version produced
/// by Excel for encrypted .xlsx files.
/// 
/// Zero dependencies. Pure C# binary I/O.
/// </summary>
internal class OleStorage : IDisposable
{
    // ── CFB constants ──
    private const ulong Magic = 0xE11AB1A1E011CFD0;
    private const int HeaderSize = 512;
    private const int SectorSize = 512;
    private const int DirEntrySize = 128;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;

    // Stream types
    private const byte StgtyStream = 2;
    private const byte StgtyStorage = 1;
    private const byte StgtyRoot = 5;

    // ── State ──
    private readonly List<uint> _fat = new();
    private readonly List<DirectoryEntry> _dirEntries = new();

    private bool _disposed;

    // ── Public structure ──
    public IReadOnlyDictionary<string, OleStreamInfo> Streams { get; }

    private readonly Dictionary<string, OleStreamInfo> _streams = new();

    /// <summary>
    /// Read an OLE2 compound file.
    /// </summary>
    public static OleStorage Read(Stream file)
    {
        var storage = new OleStorage();
        storage.ReadInternal(file);
        return storage;
    }

    /// <summary>
    /// Write an OLE2 compound file with the given streams.
    /// Each entry: (name, data bytes).
    /// </summary>
    public static void Write(Stream output, params (string Name, byte[] Data)[] streams)
    {
        WriteInternal(output, streams);
    }

    private OleStorage()
    {
        Streams = _streams;
    }

    // ═══════════════════════════════════════════════
    //  READING
    // ═══════════════════════════════════════════════

    private void ReadInternal(Stream file)
    {
        var header = new byte[HeaderSize];
        ReadExact(file, header, 0, HeaderSize);

        var magic = BitConverter.ToUInt64(header, 0);
        if (magic != Magic)
            throw new InvalidDataException("Not a valid OLE2 compound file.");

        // Header fields (v3 layout)
        // byte 22: minor version (0x3E = v3)
        // byte 23: major version (3 or 4)
        var majorVersion = header[23];
        if (majorVersion != 3)
            throw new NotSupportedException(
                $"OLE2 version {majorVersion} is not supported. Only v3 (512-byte sectors).");

        // We only support v3 (512-byte sectors)
        // byte 30-33: number of FAT sectors
        var numFatSectors = BitConverter.ToUInt32(header, 28);
        // byte 32-35: first directory sector
        var dirStartSector = BitConverter.ToUInt32(header, 30);
        // byte 44-47: first DIFAT sector
        var difatStart = BitConverter.ToUInt32(header, 40);
        // byte 48-51: first mini FAT sector
        // byte 52-55: number of mini FAT sectors
        // byte 56-59: first DIFAT sector (redundant with 44-47)
        // byte 60-63: number of DIFAT sectors

        // Read 109 initial DIFAT entries from header (offset 76)
        var difat = new List<uint>();
        for (int i = 0; i < 109; i++)
        {
            var val = BitConverter.ToUInt32(header, 76 + i * 4);
            if (val != FreeSector)
                difat.Add(val);
        }

        // Follow DIFAT chain
        var difatSector = difatStart;
        while (difatSector != EndOfChain)
        {
            var sectorBytes = new byte[SectorSize];
            ReadSector(file, difatSector, sectorBytes);
            for (int i = 0; i < 127; i++)
            {
                var val = BitConverter.ToUInt32(sectorBytes, i * 4);
                if (val != FreeSector)
                    difat.Add(val);
            }
            difatSector = BitConverter.ToUInt32(sectorBytes, 508); // last 4 bytes = next DIFAT
        }

        // Read FAT from DIFAT-referenced sectors
        _fat.Clear();
        foreach (var fatSector in difat)
        {
            var sectorBytes = new byte[SectorSize];
            ReadSector(file, fatSector, sectorBytes);
            for (int i = 0; i < 128; i++)
            {
                _fat.Add(BitConverter.ToUInt32(sectorBytes, i * 4));
            }
        }

        // Read directory
        var dirSector = dirStartSector;
        var dirBytes = new List<byte>();
        while (dirSector != EndOfChain)
        {
            var sectorBytes = new byte[SectorSize];
            ReadSector(file, dirSector, sectorBytes);
            dirBytes.AddRange(sectorBytes);
            if ((uint)_fat.Count <= dirSector) break;
            dirSector = _fat[(int)dirSector];
        }

        // Parse directory entries
        var numDirEntries = dirBytes.Count / DirEntrySize;
        for (int i = 0; i < numDirEntries; i++)
        {
            var offset = i * DirEntrySize;
            var nameLength = BitConverter.ToUInt16(dirBytes.ToArray(), offset + 64); // should not be 0 for used entries
            if (nameLength == 0) continue;

            var name = System.Text.Encoding.Unicode.GetString(
                dirBytes.ToArray(), offset, Math.Min((int)nameLength, 64));
            // Name includes null terminator in length — trim
            name = name.TrimEnd('\0');

            var entryType = dirBytes[offset + 66];
            var startSector = BitConverter.ToUInt32(dirBytes.ToArray(), offset + 116);
            var streamSize = BitConverter.ToUInt32(dirBytes.ToArray(), offset + 120);

            // First entry is root storage
            if (i == 0) continue;

            // Only care about streams (not storages)
            if (entryType == StgtyStream && startSector != EndOfChain && streamSize > 0)
            {
                _streams[name] = new OleStreamInfo(startSector, streamSize, file);
            }

            _dirEntries.Add(new DirectoryEntry(name, entryType, startSector, streamSize));
        }
    }

    // ═══════════════════════════════════════════════
    //  WRITING
    // ═══════════════════════════════════════════════

    private static void WriteInternal(Stream output, (string Name, byte[] Data)[] streams)
    {
        // Pre-calculate layout
        // For small files (< ~6.7 MB), a single FAT sector holds all chains
        // We'll support up to 127 data sectors (64 KB per stream for a couple streams)
        // which fits in one FAT + one DIFAT sector

        var entries = streams.Select(s => new StreamWriteInfo(s.Name, s.Data)).ToList();

        // Build directory: root + one entry per stream
        // Directory entries per sector: 512 / 128 = 4
        var dirEntryCount = 1 + entries.Count; // root + streams
        var dirSectorCount = (dirEntryCount + 3) / 4; // ceiling division

        // Assign sectors: header(0), FAT(count?), directory(...), data(...)
        var headerSector = 0u; // header doesn't consume a FAT sector
        var fatSectorCount = 1u; // one FAT sector gives us 128 entries
        var fatStart = 1u;
        var dirSectors = Enumerable.Range((int)fatStart + (int)fatSectorCount, dirSectorCount)
            .Select(i => (uint)i).ToList();
        var dataStart = (uint)(fatStart + fatSectorCount + dirSectorCount);

        // Allocate data sectors for each stream
        var streamAllocations = new List<(StreamWriteInfo Info, List<uint> Sectors)>();
        foreach (var entry in entries)
        {
            var sectorsNeeded = (entry.Data.Length + SectorSize - 1) / SectorSize;
            var sectors = new List<uint>();
            for (int i = 0; i < sectorsNeeded; i++)
                sectors.Add(dataStart + (uint)i);
            dataStart += (uint)sectorsNeeded;
            streamAllocations.Add((entry, sectors));
        }

        // Build FAT
        var totalSectors = (int)dataStart;
        var fat = new uint[totalSectors];
        Array.Fill(fat, FreeSector);

        // Header sector (index 0) — not tracked in FAT
        // FAT sector(s)
        fat[0] = EndOfChain; // not used (header)
        fat[1] = EndOfChain; // FAT sector — end of chain

        // Directory sectors (chained)
        for (int i = 0; i < dirSectors.Count; i++)
        {
            var idx = (int)dirSectors[i];
            fat[idx] = i < dirSectors.Count - 1 ? dirSectors[i + 1] : EndOfChain;
        }

        // Data sectors for each stream (chained)
        foreach (var (info, sectors) in streamAllocations)
        {
            for (int i = 0; i < sectors.Count; i++)
            {
                fat[(int)sectors[i]] = i < sectors.Count - 1 ? sectors[i + 1] : EndOfChain;
            }
        }

        // Build DIFAT — point to our FAT sectors
        var difat = new List<uint> { fatStart }; // just one FAT sector

        // Write header
        var header = new byte[HeaderSize];
        BitConverter.TryWriteBytes(header.AsSpan(0), Magic);
        // CLSID: 16 bytes at offset 8 — all zeros
        header[22] = 0x3E; // minor version
        header[23] = 3; // major version (v3)
        // byte order: 0xFF 0xFE = little endian
        header[24] = 0xFE;
        header[25] = 0xFF;
        // sector size exponent: 9 for 512 bytes
        header[26] = 9;
        // mini sector size exponent: 6 for 64 bytes
        header[27] = 6;
        // reserved (28-29): 0
        BitConverter.TryWriteBytes(header.AsSpan(28), (uint)fatSectorCount);
        BitConverter.TryWriteBytes(header.AsSpan(30), dirSectors[0]); // first directory sector
        // byte 32-35: transaction signature number
        // byte 36-39: mini stream cutoff size (4096)
        BitConverter.TryWriteBytes(header.AsSpan(36), 4096u);
        BitConverter.TryWriteBytes(header.AsSpan(40), dirSectors[0]); // first mini FAT sector (reuse?)
        // Actually: byte 40 = first mini FAT sector (we don't use mini streams)
        // byte 44 = first DIFAT sector (0 if none)
        // byte 48-51 = number of DIFAT sectors
        // byte 52-55 = number of mini FAT sectors
        // We'll set them all to 0/EndOfChain as needed
        BitConverter.TryWriteBytes(header.AsSpan(40), EndOfChain); // no mini FAT
        BitConverter.TryWriteBytes(header.AsSpan(44), EndOfChain); // no DIFAT (all in header)
        BitConverter.TryWriteBytes(header.AsSpan(48), 0u); // number of DIFAT sectors
        BitConverter.TryWriteBytes(header.AsSpan(52), 0u); // number of mini FAT sectors

        // DIFAT entries (offset 76, 109 entries)
        for (int i = 0; i < 109; i++)
        {
            var val = i < difat.Count ? difat[i] : FreeSector;
            BitConverter.TryWriteBytes(header.AsSpan(76 + i * 4), val);
        }

        output.Write(header, 0, HeaderSize);

        // Write FAT sector(s)
        var fatBytes = new byte[SectorSize];
        for (int i = 0; i < Math.Min(fat.Length, 128); i++)
        {
            BitConverter.TryWriteBytes(fatBytes.AsSpan(i * 4), fat[i]);
        }
        output.Write(fatBytes, 0, SectorSize);

        // Write directory sectors
        var dirBytes = new byte[dirSectorCount * SectorSize];

        // Root entry (index 0)
        WriteDirectoryEntry(dirBytes, 0, "Root Entry", StgtyRoot, EndOfChain, 0);

        // Stream entries
        for (int i = 0; i < entries.Count; i++)
        {
            var sectors = streamAllocations[i].Sectors;
            var startSector = sectors.Count > 0 ? sectors[0] : EndOfChain;
            WriteDirectoryEntry(dirBytes, (i + 1) * DirEntrySize,
                entries[i].Name, StgtyStream, startSector, entries[i].Data.Length);
        }

        output.Write(dirBytes, 0, dirBytes.Length);

        // Write data sectors
        foreach (var (info, sectors) in streamAllocations)
        {
            for (int i = 0; i < sectors.Count; i++)
            {
                var offset = i * SectorSize;
                var remaining = info.Data.Length - offset;
                var chunkSize = Math.Min(remaining, SectorSize);
                output.Write(info.Data, offset, chunkSize);
                // Pad sector
                var padding = new byte[SectorSize - chunkSize];
                output.Write(padding, 0, padding.Length);
            }
        }
    }

    private static void WriteDirectoryEntry(byte[] buffer, int offset, string name,
        byte entryType, uint startSector, int streamSize)
    {
        // Name (UTF-16LE, max 32 chars = 64 bytes including null terminator)
        var utf16Name = System.Text.Encoding.Unicode.GetBytes(name + '\0');
        if (utf16Name.Length > 64)
            throw new ArgumentException($"Stream name too long: {name}");
        Array.Copy(utf16Name, 0, buffer, offset, utf16Name.Length);
        // Name length (bytes including terminator)
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 64), (ushort)utf16Name.Length);
        // Entry type
        buffer[offset + 66] = entryType;
        // Node color: 1 = black (root), 0 = red (others)
        buffer[offset + 67] = (byte)(entryType == StgtyRoot ? 1 : 0);
        // Left sibling DIF, Right sibling DIF, child DIF: all 0xFFFFFFFF
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 68), 0xFFFFFFFFu);
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 72), 0xFFFFFFFFu);
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 76), 0xFFFFFFFFu);
        // CLSID (16 bytes at offset 80) — all zeros
        // State bits (4 bytes at 96)
        // Created/modified timestamps (8+8 bytes at 100, 108)
        // Start sector
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 116), startSector);
        // Stream size (low 32 bits)
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 120), (uint)streamSize);
        // Stream size (high 32 bits)
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 124), 0u);
    }

    // ═══════════════════════════════════════════════
    //  HELPERS
    // ═══════════════════════════════════════════════

    private void ReadSector(Stream file, uint sector, byte[] buffer)
    {
        var position = (long)(sector + 1) * SectorSize;
        file.Seek(position, SeekOrigin.Begin);
        ReadExact(file, buffer, 0, SectorSize);
    }

    private static void ReadExact(Stream stream, byte[] buffer, int offset, int count)
    {
        var total = 0;
        while (total < count)
        {
            var read = stream.Read(buffer, offset + total, count - total);
            if (read == 0)
                throw new EndOfStreamException("Unexpected end of OLE2 compound file.");
            total += read;
        }
    }

    /// <summary>
    /// Read the full content of a named stream.
    /// </summary>
    public byte[] ReadStream(string name)
    {
        if (!_streams.TryGetValue(name, out var info))
            throw new KeyNotFoundException($"Stream '{name}' not found in OLE2 container.");

        return ReadStreamData(info.StartSector, info.Size, info.File);
    }

    private byte[] ReadStreamData(uint startSector, uint size, Stream file)
    {
        var result = new byte[size];
        var offset = 0;
        var sector = startSector;

        while (sector != EndOfChain && offset < size)
        {
            var sectorBytes = new byte[SectorSize];
            ReadSector(file, sector, sectorBytes);
            var toCopy = (int)Math.Min(SectorSize, size - offset);
            Array.Copy(sectorBytes, 0, result, offset, toCopy);
            offset += toCopy;

            if ((uint)_fat.Count <= sector) break;
            sector = _fat[(int)sector];
        }

        return result;
    }

    public void Dispose()
    {
        _disposed = true;
    }

    // ── Inner types ──

    internal sealed class OleStreamInfo
    {
        public uint StartSector { get; }
        public uint Size { get; }
        internal Stream File { get; }

        internal OleStreamInfo(uint startSector, uint size, Stream file)
        {
            StartSector = startSector;
            Size = size;
            File = file;
        }
    }

    private sealed record DirectoryEntry(
        string Name, byte EntryType, uint StartSector, uint StreamSize);

    private sealed record StreamWriteInfo(string Name, byte[] Data);

    // These are needed for DIFAT reading but we use List<uint> which doesn't need explicit methods
    // (removed unused private helper methods)
}
