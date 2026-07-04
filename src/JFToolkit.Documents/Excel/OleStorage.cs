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
    private const int HeaderSize = 512;
    private const int SectorSize = 512;
    private const int DirEntrySize = 128;
    private const ulong Magic = 0xE11AB1A1E011CFD0;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;
    private const uint FatSector = 0xFFFFFFFD;
    private const uint DifSector = 0xFFFFFFFC;
    private const uint NoStream = 0xFFFFFFFF; // red-black tree null

    // ── Directory entry types ──
    private const byte StgtyEmpty = 0;
    private const byte StgtyStorage = 1;
    private const byte StgtyStream = 2;
    private const byte StgtyRoot = 5;

    // ── Red-black tree colors ──
    private const byte ColorRed = 0;
    private const byte ColorBlack = 1;

    private readonly Dictionary<string, OleStreamInfo> _streams = new();
    private readonly List<uint> _fat = new();
    private readonly List<uint> _miniFat = new();
    private readonly List<DirectoryEntry> _dirEntries = new();
    private byte[] _miniStream = Array.Empty<byte>();
    private uint _miniStreamCutoff = 4096;

    // Metadata streams required by ECMA-376 encrypted OLE2 containers.
    // These tell Excel what kind of encryption the file uses.
    public static readonly byte[] VersionStream = BuildVersionStream();
    public static readonly byte[] DataSpaceMapStream = BuildDataSpaceMapStream();
    public static readonly byte[] StrongEncryptionDataSpaceStream = BuildStrongEncryptionDataSpaceStream();
    public static readonly byte[] PrimaryStream = BuildPrimaryStream();

    public void Dispose()
    {
    }
    public IReadOnlyDictionary<string, OleStreamInfo> Streams { get; }

    // ── Public API ──

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
    /// Write an OLE2 compound file with the given streams + ECMA-376 metadata.
    /// </summary>
    public static void Write(Stream output, params (string Name, byte[] Data)[] userStreams)
    {
        // Tree structure matching msoffcrypto-tool / herumi/msoffice reference:
        //   Root (RED) → child → EncryptionInfo (BLACK)
        //     EncryptionInfo → left → DataSpaces (RED), right → EncryptedPackage (RED)
        //       DataSpaces (RED) → child → DataSpaceMap (BLACK)
        //         DataSpaceMap → left → Version (BLACK), right → DataSpaceInfo (BLACK)
        //           DataSpaceInfo → right → TransformInfo (RED), child → StrongEncryptionDataSpace (BLACK)
        //             TransformInfo → child → StrongEncryptionTransform (BLACK)
        //               StrongEncryptionTransform → child → Primary (BLACK)
        var allStreams = new List<(string Name, byte[] Data, byte Type, uint LeftSibling, uint RightSibling, uint ChildId, byte Color)>();

        var encPkg = userStreams.FirstOrDefault(s => s.Name == "EncryptedPackage");
        var encInfo = userStreams.FirstOrDefault(s => s.Name == "EncryptionInfo");

        // 0: Root Entry — RED, child→EncryptionInfo(10)
        allStreams.Add(("Root Entry", Array.Empty<byte>(), StgtyRoot, NoStream, NoStream, 10, ColorRed));

        // 1: EncryptedPackage — RED, right sibling of EncryptionInfo
        allStreams.Add(("EncryptedPackage", encPkg.Data, StgtyStream, NoStream, NoStream, NoStream, ColorRed));

        // 2: DataSpaces — RED, left sibling of EncryptionInfo
        allStreams.Add(("\x06DataSpaces", Array.Empty<byte>(), StgtyStorage, NoStream, NoStream, 4, ColorRed));

        // 3: Version — BLACK, left sibling of DataSpaceMap
        allStreams.Add(("Version", VersionStream, StgtyStream, NoStream, NoStream, NoStream, ColorBlack));

        // 4: DataSpaceMap — BLACK, child of DataSpaces, left→Version, right→DataSpaceInfo
        allStreams.Add(("DataSpaceMap", DataSpaceMapStream, StgtyStream, 3, 5, NoStream, ColorBlack));

        // 5: DataSpaceInfo — BLACK, right sibling of DataSpaceMap, right→TransformInfo, child→StrongEncryptionDataSpace
        allStreams.Add(("DataSpaceInfo", Array.Empty<byte>(), StgtyStorage, NoStream, 7, 6, ColorBlack));

        // 6: StrongEncryptionDataSpace — BLACK, child of DataSpaceInfo
        allStreams.Add(("StrongEncryptionDataSpace", StrongEncryptionDataSpaceStream, StgtyStream, NoStream, NoStream, NoStream, ColorBlack));

        // 7: TransformInfo — RED, right sibling of DataSpaceInfo, child→StrongEncryptionTransform
        allStreams.Add(("TransformInfo", Array.Empty<byte>(), StgtyStorage, NoStream, NoStream, 8, ColorRed));

        // 8: StrongEncryptionTransform — BLACK, child of TransformInfo
        allStreams.Add(("StrongEncryptionTransform", Array.Empty<byte>(), StgtyStorage, NoStream, NoStream, 9, ColorBlack));

        // 9: Primary — BLACK, child of StrongEncryptionTransform
        allStreams.Add(("\x06Primary", PrimaryStream, StgtyStream, NoStream, NoStream, NoStream, ColorBlack));

        // 10: EncryptionInfo — BLACK, child of Root, left→DataSpaces, right→EncryptedPackage
        allStreams.Add(("EncryptionInfo", encInfo.Data, StgtyStream, 2, 1, NoStream, ColorBlack));

        WriteInternal(output, allStreams);
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

        var majorVersion = header[26];
        if (majorVersion != 3)
            throw new NotSupportedException(
                $"OLE2 version {majorVersion} is not supported. Only v3 (512-byte sectors).");

        var numFatSectors = BitConverter.ToUInt32(header, 44);
        var dirStartSector = BitConverter.ToUInt32(header, 48);
        var difatStart = BitConverter.ToUInt32(header, 68);

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
            difatSector = BitConverter.ToUInt32(sectorBytes, 508);
        }

        // Read FAT
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

        // Read mini-FAT and mini-stream
        _miniStreamCutoff = BitConverter.ToUInt32(header, 56);
        var firstMiniFatSector = BitConverter.ToUInt32(header, 60);
        var numMiniFatSectors = BitConverter.ToUInt32(header, 64);

        _miniFat.Clear();
        if (firstMiniFatSector != EndOfChain && numMiniFatSectors > 0)
        {
            var mfatSector = firstMiniFatSector;
            for (int s = 0; s < numMiniFatSectors && mfatSector != EndOfChain; s++)
            {
                var sectorBytes = new byte[SectorSize];
                ReadSector(file, mfatSector, sectorBytes);
                for (int i = 0; i < 128; i++)
                    _miniFat.Add(BitConverter.ToUInt32(sectorBytes, i * 4));
                mfatSector = (uint)_fat.Count > mfatSector ? _fat[(int)mfatSector] : EndOfChain;
            }
        }

        // Read mini-stream (Root Entry's data) — done after directory parsing
        _miniStream = Array.Empty<byte>();

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
            var nameLength = BitConverter.ToUInt16(dirBytes.ToArray(), offset + 64);
            if (nameLength == 0) continue;

            var name = System.Text.Encoding.Unicode.GetString(
                dirBytes.ToArray(), offset, Math.Min((int)nameLength, 64));
            name = name.TrimEnd('\0');

            var entryType = dirBytes[offset + 66];
            var startSector = BitConverter.ToUInt32(dirBytes.ToArray(), offset + 116);
            var streamSize = BitConverter.ToUInt32(dirBytes.ToArray(), offset + 120);

            if (i == 0) continue;

            if (entryType == StgtyStream && startSector != EndOfChain && streamSize > 0)
            {
                _streams[name] = new OleStreamInfo(startSector, streamSize, file);
            }

            _dirEntries.Add(new DirectoryEntry(name, entryType, startSector, streamSize));
        }

        // If Root Entry (index 0) has data, load mini-stream
        if (numDirEntries > 0)
        {
            var rootOffset = 0 * DirEntrySize;
            var rootStart = BitConverter.ToUInt32(dirBytes.ToArray(), rootOffset + 116);
            var rootSize = BitConverter.ToUInt32(dirBytes.ToArray(), rootOffset + 120);
            if (rootStart != EndOfChain && rootSize > 0)
            {
                _miniStream = ReadStreamDataRaw(file, rootStart, (int)rootSize);
            }
        }
    }

    // ═══════════════════════════════════════════════
    //  WRITING
    // ═══════════════════════════════════════════════

    private static void WriteInternal(Stream output,
        List<(string Name, byte[] Data, byte Type, uint Left, uint Right, uint Child, byte Color)> entries)
    {
        const int MiniSectorSize = 64;
        const int MiniStreamCutoff = 4096;

        // Separate streams into mini-stream (< 4096 bytes) and regular FAT (>= 4096)
        var miniStreamEntries = new List<(int Index, byte[] Data)>();
        var fatStreamEntries = new List<(int Index, byte[] Data)>();

        for (int i = 0; i < entries.Count; i++)
        {
            var e = entries[i];
            if ((e.Type == StgtyStream || e.Type == StgtyRoot) && e.Data.Length > 0)
            {
                if (e.Data.Length < MiniStreamCutoff)
                    miniStreamEntries.Add((i, e.Data));
                else
                    fatStreamEntries.Add((i, e.Data));
            }
        }

        // Compute mini-sector layout
        var miniSectorCount = 0;
        var miniAllocations = new Dictionary<int, (int Start, int Count)>();
        foreach (var (idx, data) in miniStreamEntries)
        {
            var count = (data.Length + MiniSectorSize - 1) / MiniSectorSize;
            miniAllocations[idx] = (miniSectorCount, count);
            miniSectorCount += count;
        }

        // Root Entry mini-stream: allocate regular sectors
        var miniStreamBytes = miniSectorCount * MiniSectorSize;
        var miniStreamSectors = (miniStreamBytes + SectorSize - 1) / SectorSize;

        // Build mini-FAT
        var miniFatEntries = miniSectorCount > 0 ? miniSectorCount + 1 : 0; // +1 for sentinel
        var miniFat = new uint[miniFatEntries];
        Array.Fill(miniFat, FreeSector);
        if (miniFatEntries > 0)
            miniFat[0] = EndOfChain; // sentinel: mini-FAT itself

        foreach (var (idx, (start, count)) in miniAllocations)
        {
            for (int i = 0; i < count - 1; i++)
                miniFat[start + i] = (uint)(start + i + 1);
            miniFat[start + count - 1] = EndOfChain;
        }

        var miniFatSectorCount = (miniFatEntries * 4 + SectorSize - 1) / SectorSize;
        var numMiniFatSectors = (uint)(miniFatEntries > 0 ? miniFatSectorCount : 0);

        // Regular FAT streams
        var fatStreamAllocations = new Dictionary<int, List<uint>>();
        // Sector layout:
        //   0: FAT          (1 sector)
        //   1..: Directory  (dirSectorCount)
        //   next: Mini-FAT sectors
        //   next: Root mini-stream sectors
        //   next: Regular FAT stream sectors

        var dirSectorCount = Math.Max(1, (entries.Count + 3) / 4);
        var dirSectors = Enumerable.Range(1, dirSectorCount).Select(i => (uint)i).ToList();
        uint nextSector = (uint)(1 + dirSectorCount + miniFatSectorCount + miniStreamSectors);

        foreach (var (idx, data) in fatStreamEntries)
        {
            var sectorsNeeded = (data.Length + SectorSize - 1) / SectorSize;
            var sectors = new List<uint>();
            for (int i = 0; i < sectorsNeeded; i++)
                sectors.Add(nextSector + (uint)i);
            nextSector += (uint)sectorsNeeded;
            fatStreamAllocations[idx] = sectors;
        }

        // Build FAT
        var totalSectors = (int)nextSector;
        var fat = new uint[totalSectors];
        Array.Fill(fat, FreeSector);
        fat[0] = EndOfChain; // FAT sector

        // Directory sectors chain
        for (int i = 0; i < dirSectors.Count; i++)
            fat[(int)dirSectors[i]] = (uint)(i < dirSectors.Count - 1 ? dirSectors[i + 1] : EndOfChain);

        // Mini-FAT sectors chain
        for (int i = 0; i < miniFatSectorCount; i++)
        {
            var sec = (uint)(1 + dirSectorCount + i);
            fat[sec] = (uint)(i < miniFatSectorCount - 1 ? sec + 1 : EndOfChain);
        }

        // Root mini-stream sectors chain
        for (int i = 0; i < miniStreamSectors; i++)
        {
            var sec = (uint)(1 + dirSectorCount + miniFatSectorCount + i);
            fat[sec] = (uint)(i < miniStreamSectors - 1 ? sec + 1 : EndOfChain);
        }

        // FAT stream sectors chain
        foreach (var (_, sectors) in fatStreamAllocations)
        {
            for (int i = 0; i < sectors.Count; i++)
                fat[(int)sectors[i]] = (uint)(i < sectors.Count - 1 ? sectors[i + 1] : EndOfChain);
        }

        var difat = new List<uint> { 0 };
        uint firstMiniFatSector = miniFatEntries > 0 ? (uint)(1 + dirSectorCount) : EndOfChain;
        uint firstMiniStreamSector = miniStreamSectors > 0 ? (uint)(1 + dirSectorCount + miniFatSectorCount) : EndOfChain;

        // Write header
        var header = new byte[HeaderSize];
        BitConverter.TryWriteBytes(header.AsSpan(0), Magic);
        header[24] = 0x3E; header[25] = 0x00; // minor version
        header[26] = 0x03; header[27] = 0x00; // major version
        header[28] = 0xFE; header[29] = 0xFF; // byte order LE
        header[30] = 0x09; header[31] = 0x00; // sector size 512
        header[32] = 0x06; header[33] = 0x00; // mini sector size 64
        BitConverter.TryWriteBytes(header.AsSpan(44), (uint)1); // num FAT sectors
        BitConverter.TryWriteBytes(header.AsSpan(48), dirSectors[0]); // first dir sector
        BitConverter.TryWriteBytes(header.AsSpan(52), 0u); // transaction sig
        BitConverter.TryWriteBytes(header.AsSpan(56), (uint)MiniStreamCutoff); // mini stream cutoff
        BitConverter.TryWriteBytes(header.AsSpan(60), firstMiniFatSector); // first mini FAT
        BitConverter.TryWriteBytes(header.AsSpan(64), numMiniFatSectors); // num mini FAT
        BitConverter.TryWriteBytes(header.AsSpan(68), EndOfChain); // no DIFAT
        BitConverter.TryWriteBytes(header.AsSpan(72), 0u); // num DIFAT

        for (int i = 0; i < 109; i++)
            BitConverter.TryWriteBytes(header.AsSpan(76 + i * 4), i < difat.Count ? difat[i] : FreeSector);

        output.Write(header, 0, HeaderSize);

        // Write FAT
        var fatBytes = new byte[SectorSize];
        for (int i = 0; i < Math.Min(fat.Length, SectorSize / 4); i++)
            BitConverter.TryWriteBytes(fatBytes.AsSpan(i * 4), fat[i]);
        output.Write(fatBytes, 0, SectorSize);

        // Write directory
        var dirBytes = new byte[dirSectorCount * SectorSize];
        for (int i = 0; i < entries.Count; i++)
        {
            var e = entries[i];
            uint startSector;
            int streamSize = e.Data.Length;

            if (e.Type == StgtyRoot)
            {
                // Root Entry: start of mini-stream (regular sectors), size = mini-stream size
                startSector = firstMiniStreamSector;
                streamSize = miniStreamBytes;
            }
            else if (miniAllocations.TryGetValue(i, out var miniAlloc))
            {
                // Mini-stream: StartSector is mini-sector index, size unchanged
                startSector = (uint)miniAlloc.Start;
                streamSize = e.Data.Length;
            }
            else if (fatStreamAllocations.TryGetValue(i, out var fatSecs) && fatSecs.Count > 0)
            {
                startSector = fatSecs[0];
            }
            else
            {
                startSector = EndOfChain;
            }

            WriteDirectoryEntry(dirBytes, i * DirEntrySize,
                e.Name, e.Type, e.Left, e.Right, e.Child, startSector, streamSize, e.Color);
        }
        output.Write(dirBytes, 0, dirBytes.Length);

        // Write mini-FAT sectors
        if (miniFatEntries > 0)
        {
            var mfatBytes = new byte[miniFatSectorCount * SectorSize];
            for (int i = 0; i < miniFatEntries; i++)
                BitConverter.TryWriteBytes(mfatBytes.AsSpan(i * 4), miniFat[i]);
            output.Write(mfatBytes, 0, mfatBytes.Length);
        }

        // Write Root mini-stream sectors
        if (miniStreamBytes > 0)
        {
            var msBytes = new byte[miniStreamSectors * SectorSize];
            foreach (var (idx, data) in miniStreamEntries)
            {
                var (start, _) = miniAllocations[idx];
                Buffer.BlockCopy(data, 0, msBytes, start * MiniSectorSize, data.Length);
            }
            output.Write(msBytes, 0, msBytes.Length);
        }

        // Write FAT stream sectors
        foreach (var (idx, sectors) in fatStreamAllocations)
        {
            var data = entries[idx].Data;
            for (int i = 0; i < sectors.Count; i++)
            {
                var off = i * SectorSize;
                var remaining = data.Length - off;
                var chunkSize = Math.Min(remaining, SectorSize);
                output.Write(data, off, chunkSize);
                if (chunkSize < SectorSize)
                    output.Write(new byte[SectorSize - chunkSize], 0, SectorSize - chunkSize);
            }
        }
    }

    private static void WriteDirectoryEntry(byte[] buffer, int offset, string name,
        byte entryType, uint leftSibling, uint rightSibling, uint childId,
        uint startSector, int streamSize, byte color)
    {
        var utf16Name = System.Text.Encoding.Unicode.GetBytes(name + '\0');
        if (utf16Name.Length > 64)
            throw new ArgumentException($"Stream name too long: {name}");
        Array.Copy(utf16Name, 0, buffer, offset, utf16Name.Length);
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 64), (ushort)utf16Name.Length);
        buffer[offset + 66] = entryType;
        buffer[offset + 67] = color;
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 68), leftSibling);
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 72), rightSibling);
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 76), childId);
        // CLSID (16 bytes) — zeros
        // State bits (4 bytes) — zeros
        // Created/modified (8+8 bytes) — zeros
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 116), startSector);
        BitConverter.TryWriteBytes(buffer.AsSpan(offset + 120), (uint)streamSize);
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

    public byte[] ReadStream(string name)
    {
        if (!_streams.TryGetValue(name, out var info))
            throw new KeyNotFoundException($"Stream '{name}' not found in OLE2 container.");

        return ReadStreamData(info.StartSector, info.Size, info.File);
    }

    /// <summary>
    /// Read stream data using regular FAT sectors (for Root Entry mini-stream).
    /// </summary>
    private byte[] ReadStreamDataRaw(Stream file, uint startSector, int size)
    {
        var result = new byte[size];
        var offset = 0;
        var sector = startSector;

        while (sector != EndOfChain && sector != FreeSector && offset < size)
        {
            var sectorBytes = new byte[SectorSize];
            ReadSector(file, sector, sectorBytes);
            var toCopy = Math.Min(SectorSize, size - offset);
            Array.Copy(sectorBytes, 0, result, offset, toCopy);
            offset += toCopy;

            if ((uint)_fat.Count <= sector) break;
            sector = _fat[(int)sector];
        }

        return result;
    }

    /// <summary>
    /// Read stream data, using mini-FAT if the stream qualifies.
    /// </summary>
    private byte[] ReadStreamData(uint startSector, uint size, Stream file)
    {
        // Determine if this is a mini-stream entry
        if (_miniStream.Length > 0 && size < _miniStreamCutoff 
            && startSector != EndOfChain && startSector != FreeSector
            && startSector < _miniFat.Count)
        {
            const int MiniSectorSize = 64;
            var result = new byte[size];
            var offset = 0;
            var miniSector = startSector;

            while (miniSector != EndOfChain && miniSector != FreeSector && offset < size)
            {
                var srcOffset = (int)(miniSector * MiniSectorSize);
                if (srcOffset + MiniSectorSize > _miniStream.Length) break;
                var toCopy = (int)Math.Min(MiniSectorSize, size - offset);
                Buffer.BlockCopy(_miniStream, srcOffset, result, offset, toCopy);
                offset += toCopy;

                if (miniSector >= (uint)_miniFat.Count) break;
                miniSector = _miniFat[(int)miniSector];
            }

            return result;
        }

        // Regular FAT stream
        return ReadStreamDataRaw(file, startSector, (int)size);
    }

    // Old method removed — replaced by ReadStreamData and ReadStreamDataRaw above

    // ── Metadata stream builders ──

    private static byte[] BuildVersionStream()
    {
        // \x3c\x00\x00\x00 = 60 bytes follow
        // "Microsoft.Container.DataSpaces" in UTF-16LE + version info
        var name = "Microsoft.Container.DataSpaces";
        var nameBytes = System.Text.Encoding.Unicode.GetBytes(name);
        var result = new byte[4 + nameBytes.Length + 12];
        BitConverter.TryWriteBytes(result.AsSpan(0), nameBytes.Length + 12);
        Buffer.BlockCopy(nameBytes, 0, result, 4, nameBytes.Length);
        // Version bytes: 01 00 00 00 01 00 00 00 01 00 00 00
        result[4 + nameBytes.Length + 0] = 1;
        result[4 + nameBytes.Length + 4] = 1;
        result[4 + nameBytes.Length + 8] = 1;
        return result;
    }

    private static byte[] BuildDataSpaceMapStream()
    {
        // Maps "EncryptedPackage" → "StrongEncryptionDataSpace"
        var header = new byte[] {
            0x08, 0x00, 0x00, 0x00, // 8 bytes follow?
            0x01, 0x00, 0x00, 0x00, // 1 entry
            0x68, 0x00, 0x00, 0x00, // ref length
            0x01, 0x00, 0x00, 0x00, // ?
            0x00, 0x00, 0x00, 0x00, // ?
            0x20, 0x00, 0x00, 0x00, // name length?
        };
        var entry1 = System.Text.Encoding.Unicode.GetBytes("EncryptedPackage");
        var entry2 = System.Text.Encoding.Unicode.GetBytes("StrongEncryptionDataSpace");
        var result = new byte[header.Length + entry1.Length + 2 + entry2.Length + 2];
        Buffer.BlockCopy(header, 0, result, 0, header.Length);
        Buffer.BlockCopy(entry1, 0, result, header.Length, entry1.Length);
        Buffer.BlockCopy(entry2, 0, result, header.Length + entry1.Length + 2, entry2.Length);
        return result;
    }

    private static byte[] BuildStrongEncryptionDataSpaceStream()
    {
        var header = new byte[] {
            0x08, 0x00, 0x00, 0x00,
            0x01, 0x00, 0x00, 0x00,
            0x32, 0x00, 0x00, 0x00,
        };
        var transform = System.Text.Encoding.Unicode.GetBytes("StrongEncryptionTransform");
        var result = new byte[header.Length + transform.Length + 2];
        Buffer.BlockCopy(header, 0, result, 0, header.Length);
        Buffer.BlockCopy(transform, 0, result, header.Length, transform.Length);
        return result;
    }

    private static byte[] BuildPrimaryStream()
    {
        // CLSID + "Microsoft.Container.EncryptionTransform" + flags
        var guid = Guid.Parse("FF9A3F03-56EF-4613-BDD5-5A41C1D07246");
        var guidBytes = guid.ToByteArray();
        var name = "Microsoft.Container.EncryptionTransform";
        var nameBytes = System.Text.Encoding.Unicode.GetBytes(name);
        
        var result = new byte[4 + guidBytes.Length + 4 + nameBytes.Length + 8];
        var offset = 0;
        // Size prefix for GUID section
        BitConverter.TryWriteBytes(result.AsSpan(offset), guidBytes.Length + 4);
        offset += 4;
        Buffer.BlockCopy(guidBytes, 0, result, offset, guidBytes.Length);
        offset += guidBytes.Length;
        BitConverter.TryWriteBytes(result.AsSpan(offset), nameBytes.Length);
        offset += 4;
        Buffer.BlockCopy(nameBytes, 0, result, offset, nameBytes.Length);
        offset += nameBytes.Length;
        // Flags: 01 00 00 00 01 00 00 00
        result[offset + 0] = 1;
        result[offset + 4] = 1;
        
        return result;
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
}
