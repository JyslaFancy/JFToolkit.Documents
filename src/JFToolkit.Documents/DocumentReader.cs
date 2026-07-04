using System.Reflection;
using JFToolkit.Documents.Csv;
using JFToolkit.Documents.Excel;

namespace JFToolkit.Documents;

/// <summary>
/// Unified document reader — auto-detects format from file extension.
/// Supports .xlsx, .csv, and .json files.
/// </summary>
public static class DocumentReader
{
    /// <summary>
    /// Read a document into a list of model objects.
    /// Format is auto-detected from the file extension.
    /// </summary>
    /// <typeparam name="T">Model type with a parameterless constructor.</typeparam>
    /// <param name="path">Path to the document (.xlsx, .csv, or .json).</param>
    /// <returns>List of model objects.</returns>
    public static List<T> Read<T>(string path) where T : new()
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".xlsx" => ReadExcel<T>(path),
            ".csv" => ReadCsv<T>(path),
            ".json" => Json.JsonReader.Read<T>(path),
            _ => throw new NotSupportedException(
                $"Unsupported document format: '{ext}'. Supported formats: .xlsx, .csv, .json")
        };
    }

    private static List<T> ReadExcel<T>(string path) where T : new()
    {
        using var doc = EncryptedExcelDocument.Open(path, readOnly: true);
        var sheetNames = doc.GetSheetNames();
        if (sheetNames.Length == 0) return [];

        var sheetName = sheetNames[0];
        var props = GetMappableProperties<T>();

        // Read headers from row 1
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int col = 1;
        while (true)
        {
            var headerValue = doc.GetCellValue(sheetName, 1, col)?.ToString();
            if (string.IsNullOrEmpty(headerValue)) break;
            headers[headerValue.Trim()] = col;
            col++;
        }

        // Build property → column map
        var propColumns = new List<(PropertyInfo Prop, int Col)>();
        foreach (var prop in props)
        {
            var colAttr = prop.GetCustomAttribute<ColumnAttribute>();
            var expectedName = colAttr?.Name ?? prop.Name;

            foreach (var (header, column) in headers)
            {
                if (string.Equals(header, expectedName, StringComparison.OrdinalIgnoreCase))
                {
                    propColumns.Add((prop, column));
                    break;
                }
            }
        }

        // Read data rows (row 2+)
        var result = new List<T>();
        int row = 2;
        while (true)
        {
            bool hasData = false;
            var obj = new T();

            foreach (var (prop, column) in propColumns)
            {
                var cellValue = doc.GetCellValue(sheetName, row, column);
                if (cellValue != null && !string.IsNullOrEmpty(cellValue.ToString()))
                {
                    hasData = true;
                    SetPropertyValue(obj, prop, cellValue.ToString()!);
                }
            }

            if (!hasData) break;
            result.Add(obj);
            row++;
        }

        return result;
    }

    private static List<T> ReadCsv<T>(string path) where T : new()
    {
        return CsvReader.Read<T>(path);
    }

    private static List<PropertyInfo> GetMappableProperties<T>()
    {
        return typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite)
            .ToList();
    }

    internal static void SetPropertyValue<T>(T obj, PropertyInfo prop, string value)
    {
        try
        {
            var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

            object? converted = targetType switch
            {
                Type t when t == typeof(string) => value,
                Type t when t == typeof(int) => int.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                Type t when t == typeof(long) => long.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                Type t when t == typeof(double) => double.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                Type t when t == typeof(decimal) => decimal.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                Type t when t == typeof(float) => float.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                Type t when t == typeof(bool) => bool.Parse(value),
                Type t when t == typeof(DateTime) => DateTime.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                Type t when t == typeof(Guid) => Guid.Parse(value),
                Type t when t.IsEnum => Enum.Parse(targetType, value, ignoreCase: true),
                _ => Convert.ChangeType(value, targetType, System.Globalization.CultureInfo.InvariantCulture)
            };

            prop.SetValue(obj, converted);
        }
        catch
        {
            // Skip unparseable values
        }
    }
}
