using System.Reflection;
using System.Text;

namespace JFToolkit.Documents.Csv;

/// <summary>
/// Writes model objects to CSV files.
/// </summary>
public static class CsvWriter
{
    /// <summary>
    /// Write a collection of model objects to a CSV file.
    /// Property names are used as column headers.
    /// </summary>
    public static void Write<T>(string path, IEnumerable<T> data, CsvOptions? options = null)
    {
        options ??= CsvOptions.Default;
        var items = data.ToList();
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead)
            .ToList();

        var sb = new StringBuilder();

        // Header row
        if (options.HasHeaderRow)
        {
            var headers = props.Select(p =>
            {
                var colAttr = p.GetCustomAttribute<ColumnAttribute>();
                return EscapeField(colAttr?.Name ?? p.Name, options.Delimiter);
            });
            sb.AppendLine(string.Join(options.Delimiter.ToString(), headers));
        }

        // Data rows
        foreach (var item in items)
        {
            var fields = props.Select(p =>
            {
                var value = p.GetValue(item);
                var str = value?.ToString() ?? "";
                return EscapeField(str, options.Delimiter);
            });
            sb.AppendLine(string.Join(options.Delimiter.ToString(), fields));
        }

        File.WriteAllText(path, sb.ToString(), options.Encoding);
    }

    private static string EscapeField(string field, char delimiter)
    {
        if (string.IsNullOrEmpty(field))
            return "";

        bool needsQuoting = field.Contains(delimiter)
            || field.Contains('"')
            || field.Contains('\n')
            || field.Contains('\r');

        if (!needsQuoting)
            return field;

        // Escape double quotes by doubling them
        var escaped = field.Replace("\"", "\"\"");
        return $"\"{escaped}\"";
    }
}
