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
    private readonly List<DirectoryEntry> _dirEntries = new();

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
    }

    // ═══════════════════════════════════════════════
    //  WRITING
    // ═══════════════════════════════════════════════

    private static void WriteInternal(Stream output,
        List<(string Name, byte[] Data, byte Type, uint Left, uint Right, uint Child, byte Color)> entries)
    {
        // Collect only stream entries (non-storage) for sector allocation
        var streamEntries = entries
            .Select((e, i) => (Index: i, Entry: e))
            .Where(x => x.Entry.Type == StgtyStream || x.Entry.Type == StgtyRoot)
            .ToList();

        var dirEntryCount = entries.Count;
        var dirSectorCount = (dirEntryCount + 3) / 4;

        var fatSectorCount = 1u;
        var dirSectors = Enumerable.Range(1, dirSectorCount).Select(i => (uint)i).ToList();
        var dataStart = (uint)(1 + dirSectorCount);

        // Allocate data sectors for each stream
        var streamAllocations = new Dictionary<int, List<uint>>();
        uint nextSector = dataStart;
        foreach (var (idx, entry) in streamEntries)
        {
            var sectorsNeeded = (entry.Data.Length + SectorSize - 1) / SectorSize;
            if (sectorsNeeded == 0) continue;
            var sectors = new List<uint>();
            for (int i = 0; i < sectorsNeeded; i++)
                sectors.Add(nextSector + (uint)i);
            nextSector += (uint)sectorsNeeded;
            streamAllocations[idx] = sectors;
        }

        // Build FAT
        var totalSectors = (int)nextSector;
        var fat = new uint[totalSectors];
        Array.Fill(fat, FreeSector);
        fat[0] = EndOfChain; // FAT sector

        // Directory sectors (chained)
        for (int i = 0; i < dirSectors.Count; i++)
        {
            var idx = (int)dirSectors[i];
            fat[idx] = i < dirSectors.Count - 1 ? dirSectors[i + 1] : EndOfChain;
        }

        // Data sectors for each stream (chained)
        foreach (var (entryIdx, sectors) in streamAllocations)
        {
            for (int i = 0; i < sectors.Count; i++)
            {
                fat[(int)sectors[i]] = i < sectors.Count - 1 ? sectors[i + 1] : EndOfChain;
            }
        }

        var difat = new List<uint> { 0 }; // FAT at sector 0

        // Write header
        var header = new byte[HeaderSize];
        BitConverter.TryWriteBytes(header.AsSpan(0), Magic);
        // Offset 0x18-0x19: Minor Version (after 16-byte CLSID at 0x08-0x17)
        header[24] = 0x3E; // minor version
        header[25] = 0x00;
        // Offset 0x1A-0x1B: Major Version
        header[26] = 0x03; // major version
        header[27] = 0x00;
        // Offset 0x1C-0x1D: Byte Order (0xFFFE = little endian)
        header[28] = 0xFE;
        header[29] = 0xFF;
        // Offset 0x1E-0x1F: Sector Size (9 → 2^9 = 512)
        header[30] = 0x09;
        header[31] = 0x00;
        // Offset 0x20-0x21: Mini Sector Size (6 → 2^6 = 64)
        header[32] = 0x06;
        header[33] = 0x00;
        BitConverter.TryWriteBytes(header.AsSpan(44), fatSectorCount);
        BitConverter.TryWriteBytes(header.AsSpan(48), dirSectors[0]);
        BitConverter.TryWriteBytes(header.AsSpan(52), 0u); // transaction sig (0 = finalized)
        BitConverter.TryWriteBytes(header.AsSpan(56), 4096u); // mini stream cutoff
        BitConverter.TryWriteBytes(header.AsSpan(60), EndOfChain); // first mini FAT
        BitConverter.TryWriteBytes(header.AsSpan(64), 0u);        // mini FAT sectors
        BitConverter.TryWriteBytes(header.AsSpan(68), EndOfChain); // no DIFAT chain
        BitConverter.TryWriteBytes(header.AsSpan(72), 0u);        // DIFAT sectors

        for (int i = 0; i < 109; i++)
        {
            var val = i < difat.Count ? difat[i] : FreeSector;
            BitConverter.TryWriteBytes(header.AsSpan(76 + i * 4), val);
        }

        output.Write(header, 0, HeaderSize);

        // Write FAT
        var fatBytes = new byte[SectorSize];
        for (int i = 0; i < Math.Min(fat.Length, 128); i++)
            BitConverter.TryWriteBytes(fatBytes.AsSpan(i * 4), fat[i]);
        output.Write(fatBytes, 0, SectorSize);

        // Write directory
        var dirBytes = new byte[dirSectorCount * SectorSize];
        for (int i = 0; i < entries.Count; i++)
        {
            var e = entries[i];
            uint startSector = streamAllocations.TryGetValue(i, out var secs) && secs.Count > 0
                ? secs[0] : EndOfChain;
            WriteDirectoryEntry(dirBytes, i * DirEntrySize,
                e.Name, e.Type, e.Left, e.Right, e.Child, startSector, e.Data.Length, e.Color);
        }
        output.Write(dirBytes, 0, dirBytes.Length);

        // Write data sectors
        foreach (var (idx, sectors) in streamAllocations)
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
