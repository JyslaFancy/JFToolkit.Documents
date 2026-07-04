using JFToolkit.Documents;
using Xunit;

namespace JFToolkit.Documents.Tests;

public class JsonReaderWriterTests
{
    [Fact]
    public void Read_ValidJsonArray_ReturnsData()
    {
        var json = """[{"name":"Widget","price":9.99,"stock":100},{"name":"Gadget","price":24.50,"stock":50}]""";
        var path = GetTempPath("read.json");
        File.WriteAllText(path, json);

        var result = Json.JsonReader.Read<Product>(path);

        Assert.Equal(2, result.Count);
        Assert.Equal("Widget", result[0].Name);
        Assert.Equal(9.99m, result[0].Price);
    }

    [Fact]
    public void Read_EmptyArray_ReturnsEmptyList()
    {
        var path = GetTempPath("empty.json");
        File.WriteAllText(path, "[]");

        var result = Json.JsonReader.Read<Product>(path);

        Assert.Empty(result);
    }

    [Fact]
    public void Write_Data_ProducesValidJson()
    {
        var data = new List<Product>
        {
            new() { Name = "Widget", Price = 9.99m, Stock = 100 }
        };
        var path = GetTempPath("write.json");

        Json.JsonWriter.Write(path, data);

        var json = File.ReadAllText(path);
        Assert.Contains("\"name\": \"Widget\"", json);  // camelCase
        Assert.Contains("\"price\": 9.99", json);
    }

    [Fact]
    public void Roundtrip_PreservesData()
    {
        var original = new List<Product>
        {
            new() { Name = "A", Price = 1.0m, Stock = 1 },
            new() { Name = "B", Price = 2.0m, Stock = 2 },
        };
        var path = GetTempPath("roundtrip.json");

        Json.JsonWriter.Write(path, original);
        var result = Json.JsonReader.Read<Product>(path);

        Assert.Equal(2, result.Count);
        Assert.Equal("A", result[0].Name);
        Assert.Equal(1.0m, result[0].Price);
    }

    private static string GetTempPath(string name)
    {
        var dir = Path.Combine(Path.GetTempPath(), "jftoolkit_tests");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, name);
    }

    public class Product
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public int Stock { get; set; }
    }
}
