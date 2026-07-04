using System.Text.Json;

namespace JFToolkit.Documents.Json;

/// <summary>
/// Writes model objects to JSON files.
/// </summary>
public static class JsonWriter
{
    private static readonly JsonSerializerOptions DefaultOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    /// <summary>
    /// Write a collection of model objects to a JSON file as a JSON array.
    /// </summary>
    public static void Write<T>(string path, IEnumerable<T> data)
    {
        var json = JsonSerializer.Serialize(data.ToList(), DefaultOptions);
        File.WriteAllText(path, json);
    }
}
