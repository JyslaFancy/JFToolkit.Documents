using System.Globalization;
using System.Reflection;
using System.Text;

namespace JFToolkit.Documents.Csv;

/// <summary>
/// Reads CSV files and maps rows to model objects.
/// First row is treated as header by default.
/// </summary>
public static class CsvReader
{
    /// <summary>
    /// Read a CSV file into a list of model objects.
    /// Property names are matched to column headers (case-insensitive).
    /// </summary>
    public static List<T> Read<T>(string path, CsvOptions? options = null) where T : new()
    {
        options ??= CsvOptions.Default;
        var lines = File.ReadAllLines(path, options.Encoding);

        if (lines.Length == 0)
            return [];

        var headers = ParseLine(lines[0], options.Delimiter);
        var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < headers.Length; i++)
        {
            var h = headers[i].Trim();
            if (!string.IsNullOrEmpty(h))
                headerMap[h] = i;
        }

        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite)
            .ToList();

        // Build column index → property map
        var columnMap = new PropertyInfo?[headers.Length];
        for (int i = 0; i < headers.Length; i++)
        {
            var h = headers[i].Trim();
            if (string.IsNullOrEmpty(h)) continue;

            foreach (var prop in props)
            {
                var colAttr = prop.GetCustomAttribute<ColumnAttribute>();
                var expectedName = colAttr?.Name ?? prop.Name;
                if (string.Equals(expectedName, h, StringComparison.OrdinalIgnoreCase))
                {
                    columnMap[i] = prop;
                    break;
                }
            }
        }

        var result = new List<T>();
        for (int lineIdx = 1; lineIdx < lines.Length; lineIdx++)
        {
            var line = lines[lineIdx].Trim();
            if (string.IsNullOrEmpty(line)) continue;

            var fields = ParseLine(line, options.Delimiter);
            var obj = new T();

            for (int col = 0; col < Math.Min(fields.Length, columnMap.Length); col++)
            {
                var prop = columnMap[col];
                if (prop == null) continue;

                var value = fields[col].Trim();
                if (string.IsNullOrEmpty(value)) continue;

                SetPropertyValue(obj, prop, value);
            }

            result.Add(obj);
        }

        return result;
    }

    private static string[] ParseLine(string line, char delimiter)
    {
        var fields = new List<string>();
        var current = new StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    // Check for escaped quote ""
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++; // skip next quote
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    current.Append(c);
                }
            }
            else
            {
                if (c == '"')
                {
                    inQuotes = true;
                }
                else if (c == delimiter)
                {
                    fields.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }
        }
        fields.Add(current.ToString());
        return fields.ToArray();
    }

    private static void SetPropertyValue<T>(T obj, PropertyInfo prop, string value)
    {
        try
        {
            var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

            object? converted = targetType switch
            {
                Type t when t == typeof(string) => value,
                Type t when t == typeof(int) => int.Parse(value, CultureInfo.InvariantCulture),
                Type t when t == typeof(long) => long.Parse(value, CultureInfo.InvariantCulture),
                Type t when t == typeof(double) => double.Parse(value, CultureInfo.InvariantCulture),
                Type t when t == typeof(decimal) => decimal.Parse(value, CultureInfo.InvariantCulture),
                Type t when t == typeof(float) => float.Parse(value, CultureInfo.InvariantCulture),
                Type t when t == typeof(bool) => bool.Parse(value),
                Type t when t == typeof(DateTime) => DateTime.Parse(value, CultureInfo.InvariantCulture),
                Type t when t == typeof(Guid) => Guid.Parse(value),
                _ => Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture)
            };

            prop.SetValue(obj, converted);
        }
        catch
        {
            // Skip fields that can't be parsed — leave default value
        }
    }
}

/// <summary>
/// Options for CSV reading/writing.
/// </summary>
public class CsvOptions
{
    public static readonly CsvOptions Default = new();

    /// <summary>Field delimiter character. Default: ','</summary>
    public char Delimiter { get; set; } = ',';

    /// <summary>File encoding. Default: UTF-8</summary>
    public Encoding Encoding { get; set; } = Encoding.UTF8;

    /// <summary>Include header row when writing. Default: true</summary>
    public bool HasHeaderRow { get; set; } = true;
}
