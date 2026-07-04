using System.Text.Json;

namespace JFToolkit.Documents.Json;

/// <summary>
/// Reads JSON files into model objects.
/// Expects a JSON array at the root.
/// </summary>
public static class JsonReader
{
    private static readonly JsonSerializerOptions DefaultOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = false
    };

    /// <summary>
    /// Read a JSON file (JSON array) into a list of model objects.
    /// </summary>
    public static List<T> Read<T>(string path) where T : new()
    {
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<List<T>>(json, DefaultOptions) ?? [];
    }
}
