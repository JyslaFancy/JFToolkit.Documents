using System.Reflection;
using JFToolkit.Documents.Csv;
using JFToolkit.Documents.Excel;

namespace JFToolkit.Documents;

/// <summary>
/// Unified document writer — auto-detects format from file extension.
/// Supports .xlsx, .csv, and .json files.
/// </summary>
public static class DocumentWriter
{
    /// <summary>
    /// Write a collection of model objects to a document.
    /// Format is auto-detected from the file extension.
    /// </summary>
    /// <typeparam name="T">Model type.</typeparam>
    /// <param name="path">Output path (.xlsx, .csv, or .json).</param>
    /// <param name="data">Collection of model objects to write.</param>
    public static void Write<T>(string path, IEnumerable<T> data)
    {
        var items = data.ToList();
        var ext = Path.GetExtension(path).ToLowerInvariant();

        switch (ext)
        {
            case ".xlsx":
                WriteExcel(path, items);
                break;
            case ".csv":
                CsvWriter.Write(path, items);
                break;
            case ".json":
                Json.JsonWriter.Write(path, items);
                break;
            default:
                throw new NotSupportedException(
                    $"Unsupported document format: '{ext}'. Supported formats: .xlsx, .csv, .json");
        }
    }

    private static void WriteExcel<T>(string path, List<T> data)
    {
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead)
            .ToList();

        // Create a fresh workbook
        using var doc = CreateEmptyWorkbook(path);

        var sheetNames = doc.GetSheetNames();
        var sheetName = sheetNames.Length > 0 ? sheetNames[0] : "Sheet1";

        // Write header row
        for (int col = 0; col < props.Count; col++)
        {
            var colAttr = props[col].GetCustomAttribute<ColumnAttribute>();
            var header = colAttr?.Name ?? props[col].Name;
            doc.SetCellValue(sheetName, 1, col + 1, header);
        }

        // Write data rows
        for (int row = 0; row < data.Count; row++)
        {
            var item = data[row];
            for (int col = 0; col < props.Count; col++)
            {
                var value = props[col].GetValue(item);
                if (value != null)
                {
                    doc.SetCellValue(sheetName, row + 2, col + 1, value);
                }
            }
        }

        doc.SaveAs(path, null); // Save unencrypted
    }

    /// <summary>
    /// Creates an empty .xlsx workbook from a minimal template.
    /// </summary>
    private static EncryptedExcelDocument CreateEmptyWorkbook(string path)
    {
        // Create a minimal valid .xlsx in memory
        var zipBytes = CreateMinimalXlsx();
        var tempPath = Path.Combine(Path.GetTempPath(), $"jft_empty_{Guid.NewGuid():N}.xlsx");
        File.WriteAllBytes(tempPath, zipBytes);

        try
        {
            return EncryptedExcelDocument.Open(tempPath, readOnly: false);
        }
        finally
        {
            try { File.Delete(tempPath); } catch { /* ignore */ }
        }
    }

    private static byte[] CreateMinimalXlsx()
    {
        using var ms = new System.IO.MemoryStream();
        using (var archive = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
        {
            // [Content_Types].xml
            var contentTypes = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
  <Override PartName=""/xl/sharedStrings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml""/>
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>";

            var entry = archive.CreateEntry("[Content_Types].xml", System.IO.Compression.CompressionLevel.Optimal);
            using (var writer = new System.IO.StreamWriter(entry.Open()))
                writer.Write(contentTypes);

            // _rels/.rels
            var rels = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>";

            entry = archive.CreateEntry("_rels/.rels", System.IO.Compression.CompressionLevel.Optimal);
            using (var writer = new System.IO.StreamWriter(entry.Open()))
                writer.Write(rels);

            // xl/workbook.xml
            var workbook = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets>
    <sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/>
  </sheets>
</workbook>";

            entry = archive.CreateEntry("xl/workbook.xml", System.IO.Compression.CompressionLevel.Optimal);
            using (var writer = new System.IO.StreamWriter(entry.Open()))
                writer.Write(workbook);

            // xl/_rels/workbook.xml.rels
            var wbRels = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"" Target=""sharedStrings.xml""/>
  <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>";

            entry = archive.CreateEntry("xl/_rels/workbook.xml.rels", System.IO.Compression.CompressionLevel.Optimal);
            using (var writer = new System.IO.StreamWriter(entry.Open()))
                writer.Write(wbRels);

            // xl/worksheets/sheet1.xml
            var sheet = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheetData/>
</worksheet>";

            entry = archive.CreateEntry("xl/worksheets/sheet1.xml", System.IO.Compression.CompressionLevel.Optimal);
            using (var writer = new System.IO.StreamWriter(entry.Open()))
                writer.Write(sheet);

            // xl/sharedStrings.xml
            var sst = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""0"" uniqueCount=""0""/>";

            entry = archive.CreateEntry("xl/sharedStrings.xml", System.IO.Compression.CompressionLevel.Optimal);
            using (var writer = new System.IO.StreamWriter(entry.Open()))
                writer.Write(sst);

            // xl/styles.xml
            var styles = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <fonts count=""1"">
    <font><sz val=""11""/><name val=""Calibri""/></font>
  </fonts>
  <fills count=""2"">
    <fill><patternFill patternType=""none""/></fill>
    <fill><patternFill patternType=""gray125""/></fill>
  </fills>
  <borders count=""1"">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count=""1"">
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/>
  </cellStyleXfs>
  <cellXfs count=""1"">
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/>
  </cellXfs>
</styleSheet>";

            entry = archive.CreateEntry("xl/styles.xml", System.IO.Compression.CompressionLevel.Optimal);
            using (var writer = new System.IO.StreamWriter(entry.Open()))
                writer.Write(styles);
        }

        return ms.ToArray();
    }
}
