using System.IO.Compression;
using System.Xml.Linq;

namespace JFToolkit.Documents.Excel;

/// <summary>
/// Reads and writes cell values from an unencrypted .xlsx ZIP.
/// 
/// Stores modifications in memory as a sparse dictionary keyed
/// by (sheet, row, col). On save, preserves the original structure
/// and only updates changed cells.
/// 
/// Zero dependencies — uses System.IO.Compression and System.Xml.Linq.
/// </summary>
internal class XlsxCellStore
{
    // ── State ──
    private readonly Dictionary<string, SheetData> _sheets = new();
    private readonly Dictionary<string, byte[]> _zipEntries = new();
    private readonly List<string> _sharedStrings = new();
    private bool _stringsModified;

    // ── Public API ──

    /// <summary>
    /// Load from raw ZIP bytes (the decrypted inner package).
    /// </summary>
    public static XlsxCellStore Load(byte[] zipData)
    {
        var store = new XlsxCellStore();
        store.ReadZip(zipData);
        return store;
    }

    /// <summary>
    /// Get all sheet names in the workbook.
    /// </summary>
    public string[] GetSheetNames() => _sheets.Keys.ToArray();

    /// <summary>
    /// Get the first sheet's name, or null if no sheets.
    /// </summary>
    public string? FirstSheetName => _sheets.Keys.FirstOrDefault();

    /// <summary>
    /// Number of sheets.
    /// </summary>
    public int SheetCount => _sheets.Count;

    /// <summary>
    /// Get a cell value (1-based row/col).
    /// Returns null if the cell is empty.
    /// </summary>
    public object? GetCellValue(string sheetName, int row, int col)
    {
        if (!_sheets.TryGetValue(sheetName, out var sheet))
            throw new ArgumentException($"Sheet '{sheetName}' not found.");

        return sheet.GetValue(row, col, _sharedStrings);
    }

    /// <summary>
    /// Set a cell value (1-based row/col).
    /// Accepts string, double, int, DateTime, bool.
    /// </summary>
    public void SetCellValue(string sheetName, int row, int col, object value)
    {
        if (!_sheets.TryGetValue(sheetName, out var sheet))
            throw new ArgumentException($"Sheet '{sheetName}' not found.");

        if (value is string strValue)
        {
            // Add to shared strings table
            var idx = _sharedStrings.IndexOf(strValue);
            if (idx < 0)
            {
                idx = _sharedStrings.Count;
                _sharedStrings.Add(strValue);
                _stringsModified = true;
            }
            sheet.SetValue(row, col, CellValue.SharedString(idx));
        }
        else if (value is double or int or float or decimal)
        {
            sheet.SetValue(row, col, CellValue.Number(Convert.ToDouble(value)));
        }
        else if (value is bool b)
        {
            sheet.SetValue(row, col, CellValue.Boolean(b));
        }
        else if (value is DateTime dt)
        {
            // Store as OLE date (double)
            sheet.SetValue(row, col, CellValue.Number(dt.ToOADate()));
        }
        else
        {
            // Fallback — convert to string
            SetCellValue(sheetName, row, col, value.ToString() ?? "");
        }
    }

    /// <summary>
    /// Build the modified ZIP bytes.
    /// </summary>
    public byte[] Save()
    {
        return BuildZip();
    }

    // ═══════════════════════════════════════════════
    //  ZIP READING
    // ═══════════════════════════════════════════════

    private void ReadZip(byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        // First pass: read shared strings
        var ssEntry = archive.GetEntry("xl/sharedStrings.xml");
        if (ssEntry != null)
        {
            using var stream = ssEntry.Open();
            var doc = XElement.Load(stream);
            var ns = doc.GetDefaultNamespace();
            foreach (var si in doc.Elements(ns + "si"))
            {
                var t = si.Element(ns + "t");
                _sharedStrings.Add(t?.Value ?? "");
            }
        }

        // Second pass: find all sheet files
        var sheetEntries = new List<(string Name, ZipArchiveEntry Entry)>();
        foreach (var entry in archive.Entries)
        {
            if (entry.FullName.StartsWith("xl/worksheets/sheet") &&
                entry.FullName.EndsWith(".xml"))
            {
                sheetEntries.Add((entry.FullName, entry));
            }
        }

        // Read workbook.xml to map sheet names to files
        var wbEntry = archive.GetEntry("xl/workbook.xml");
        if (wbEntry == null)
            throw new InvalidDataException("Not a valid .xlsx file: missing xl/workbook.xml.");

        var sheetNames = new Dictionary<string, string>(); // file -> name
        using (var wbStream = wbEntry.Open())
        {
            var wbDoc = XElement.Load(wbStream);
            var ns = wbDoc.GetDefaultNamespace();
            var sheets = wbDoc.Element(ns + "sheets");
            if (sheets != null)
            {
                foreach (var sheet in sheets.Elements(ns + "sheet"))
                {
                    var name = sheet.Attribute("name")?.Value ?? "";
                    var rId = sheet.Attribute("id")?.Value ?? ""; // r:id from relationships

                    // We need relationship mapping. For simplicity, match by
                    // the sheet position and xml filename.

                    // Sheet files are named sheet1.xml, sheet2.xml, etc.
                    var sheetId = sheet.Attribute("sheetId")?.Value;
                    if (sheetId != null)
                    {
                        var fileName = $"xl/worksheets/sheet{sheetId}.xml";
                        sheetNames[fileName] = name;
                    }
                }
            }
        }

        // Parse each sheet
        foreach (var (path, entry) in sheetEntries)
        {
            var name = sheetNames.TryGetValue(path, out var n) ? n : Path.GetFileNameWithoutExtension(path);

            using var stream = entry.Open();
            var doc = XElement.Load(stream);
            var ns = doc.GetDefaultNamespace();

            var sheetDataElement = doc.Element(ns + "sheetData");
            var sheetData = new Dictionary<(int Row, int Col), CellValue>();

            if (sheetDataElement != null)
            {
                foreach (var rowEl in sheetDataElement.Elements(ns + "row"))
                {
                    int row = int.Parse(rowEl.Attribute("r")?.Value ?? "0");
                    foreach (var cell in rowEl.Elements(ns + "c"))
                    {
                        var cellRef = cell.Attribute("r")?.Value ?? "";
                        var col = CellRefToColumn(cellRef);
                        var cellType = cell.Attribute("t")?.Value;
                        var value = cell.Element(ns + "v")?.Value;

                        if (value != null)
                        {
                            if (cellType == "s") // shared string
                            {
                                var idx = int.Parse(value);
                                if (idx < _sharedStrings.Count)
                                    sheetData[(row, col)] = CellValue.SharedString(idx);
                            }
                            else if (cellType == "b") // boolean
                            {
                                sheetData[(row, col)] = CellValue.Boolean(value == "1");
                            }
                            else // number or date
                            {
                                sheetData[(row, col)] = CellValue.Number(double.Parse(value,
                                    System.Globalization.CultureInfo.InvariantCulture));
                            }
                        }
                    }
                }
            }

            _sheets[name] = new SheetData(name, path, sheetData);
        }

        // Store all original ZIP entries for reconstruction
        foreach (var entry in archive.Entries)
        {
            using var stream = entry.Open();
            var entryData = new byte[entry.Length];
            stream.ReadExactly(entryData, 0, entryData.Length);
            _zipEntries[entry.FullName] = entryData;
        }
    }

    // ═══════════════════════════════════════════════
    //  ZIP WRITING
    // ═══════════════════════════════════════════════

    private byte[] BuildZip()
    {
        var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            // Copy unmodified entries
            foreach (var (name, data) in _zipEntries)
            {
                // Skip sheet entries that have been modified — we'll write new versions below
                if (name.StartsWith("xl/worksheets/sheet") && name.EndsWith(".xml"))
                {
                    var sheetName = _sheets.Values
                        .FirstOrDefault(s => s.XmlFilePath == name)?.Name;
                    if (sheetName != null && _sheets.TryGetValue(sheetName, out var s) && s.IsModified)
                        continue;
                }
                if (name == "xl/sharedStrings.xml" && _stringsModified)
                    continue;

                var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
                using var es = entry.Open();
                es.Write(data, 0, data.Length);
            }

            // Write modified sheet XMLs
            foreach (var (_, sheet) in _sheets)
            {
                if (!sheet.IsModified) continue;

                var xml = sheet.BuildXml();
                var entry = archive.CreateEntry(sheet.XmlFilePath, CompressionLevel.Optimal);
                using var es = entry.Open();
                var bytes = System.Text.Encoding.UTF8.GetBytes(xml);
                es.Write(bytes, 0, bytes.Length);
            }

            // Write shared strings if modified
            if (_stringsModified)
            {
                var xml = BuildSharedStringsXml();
                var entry = archive.CreateEntry("xl/sharedStrings.xml", CompressionLevel.Optimal);
                using var es = entry.Open();
                var bytes = System.Text.Encoding.UTF8.GetBytes(xml);
                es.Write(bytes, 0, bytes.Length);
            }
        }

        return output.ToArray();
    }

    private string BuildSharedStringsXml()
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var doc = new XElement(ns + "sst",
            new XAttribute("uniqueCount", _sharedStrings.Count));

        foreach (var s in _sharedStrings)
        {
            doc.Add(new XElement(ns + "si",
                new XElement(ns + "t", new XText(s))));
        }

        return doc.ToString(SaveOptions.DisableFormatting);
    }

    // ═══════════════════════════════════════════════
    //  HELPERS
    // ═══════════════════════════════════════════════

    private static int CellRefToColumn(string cellRef)
    {
        // Extract column letters from "A1", "BC23", etc.
        var letters = new string(cellRef.TakeWhile(char.IsLetter).ToArray());
        var col = 0;
        foreach (var c in letters)
        {
            col = col * 26 + (char.ToUpper(c) - 'A' + 1);
        }
        return col;
    }

    private static string ColumnToCellRef(int col, int row)
    {
        var result = new System.Text.StringBuilder();
        while (col > 0)
        {
            col--;
            result.Insert(0, (char)('A' + col % 26));
            col /= 26;
        }
        result.Append(row);
        return result.ToString();
    }

    // ═══════════════════════════════════════════════
    //  INNER TYPES
    // ═══════════════════════════════════════════════

    private record CellValue
    {
        public enum CellType { Number, SharedString, Boolean }
        public CellType Type { get; init; }
        public double NumericValue { get; init; }
        public int StringIndex { get; init; }
        public bool BoolValue { get; init; }

        public static CellValue Number(double v) => new() { Type = CellType.Number, NumericValue = v };
        public static CellValue SharedString(int idx) => new() { Type = CellType.SharedString, StringIndex = idx };
        public static CellValue Boolean(bool v) => new() { Type = CellType.Boolean, BoolValue = v, NumericValue = v ? 1 : 0 };
    }

    private sealed class SheetData
    {
        public string Name { get; }
        public string XmlFilePath { get; }
        public string XmlFileName => Path.GetFileName(XmlFilePath);
        public bool IsModified { get; private set; }

        private readonly Dictionary<(int Row, int Col), CellValue> _cells;

        public SheetData(string name, string xmlFilePath,
            Dictionary<(int Row, int Col), CellValue> cells)
        {
            Name = name;
            XmlFilePath = xmlFilePath;
            _cells = cells;
        }

        public object? GetValue(int row, int col, List<string> sharedStrings)
        {
            if (!_cells.TryGetValue((row, col), out var cv))
                return null;

            return cv.Type switch
            {
                CellValue.CellType.Number => cv.NumericValue,
                CellValue.CellType.SharedString when cv.StringIndex < sharedStrings.Count
                    => sharedStrings[cv.StringIndex],
                CellValue.CellType.SharedString => "",
                CellValue.CellType.Boolean => cv.BoolValue,
                _ => null
            };
        }

        public void SetValue(int row, int col, CellValue value)
        {
            _cells[(row, col)] = value;
            IsModified = true;
        }

        public string BuildXml()
        {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var doc = new XElement(ns + "worksheet");

            var sheetData = new XElement(ns + "sheetData");

            // Group cells by row
            var rows = _cells.GroupBy(kv => kv.Key.Row)
                .OrderBy(g => g.Key);

            foreach (var rowGroup in rows)
            {
                var rowEl = new XElement(ns + "row",
                    new XAttribute("r", rowGroup.Key));

                foreach (var ((_, col), cell) in rowGroup.OrderBy(kv => kv.Key.Col))
                {
                    var cellRef = ColumnToCellRef(col, rowGroup.Key);
                    var cellEl = new XElement(ns + "c",
                        new XAttribute("r", cellRef));

                    switch (cell.Type)
                    {
                        case CellValue.CellType.Number:
                            cellEl.SetAttributeValue("t", "n");
                            cellEl.Add(new XElement(ns + "v",
                                cell.NumericValue.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                            break;

                        case CellValue.CellType.SharedString:
                            cellEl.SetAttributeValue("t", "s");
                            cellEl.Add(new XElement(ns + "v", cell.StringIndex));
                            break;

                        case CellValue.CellType.Boolean:
                            cellEl.SetAttributeValue("t", "b");
                            cellEl.Add(new XElement(ns + "v", cell.BoolValue ? "1" : "0"));
                            break;
                    }

                    rowEl.Add(cellEl);
                }

                sheetData.Add(rowEl);
            }

            doc.Add(sheetData);
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                   doc.ToString(SaveOptions.DisableFormatting);
        }
    }
}
