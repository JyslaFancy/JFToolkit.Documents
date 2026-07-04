using JFToolkit.Documents;
using Xunit;

namespace JFToolkit.Documents.Tests;

public class DocumentReaderWriterTests
{
    [Fact]
    public void Read_Csv_AutoDetected()
    {
        var csv = "Name,Price,Stock\nWidget,9.99,100";
        var path = GetTempPath("test.csv");
        File.WriteAllText(path, csv);

        var result = DocumentReader.Read<Product>(path);

        Assert.Single(result);
        Assert.Equal("Widget", result[0].Name);
    }

    [Fact]
    public void Read_Json_AutoDetected()
    {
        var json = """[{"name":"Widget","price":9.99,"stock":100}]""";
        var path = GetTempPath("test.json");
        File.WriteAllText(path, json);

        var result = DocumentReader.Read<Product>(path);

        Assert.Single(result);
        Assert.Equal("Widget", result[0].Name);
    }

    [Fact]
    public void Read_UnknownExtension_Throws()
    {
        Assert.Throws<NotSupportedException>(() =>
            DocumentReader.Read<Product>("file.unknown"));
    }

    [Fact]
    public void Write_Csv_AutoDetected()
    {
        var data = new List<Product> { new() { Name = "X", Price = 1, Stock = 1 } };
        var path = GetTempPath("out.csv");

        DocumentWriter.Write(path, data);

        Assert.True(File.Exists(path));
        var lines = File.ReadAllLines(path);
        Assert.Equal("Name,Price,Stock", lines[0]);
    }

    [Fact]
    public void Write_Json_AutoDetected()
    {
        var data = new List<Product> { new() { Name = "X", Price = 1, Stock = 1 } };
        var path = GetTempPath("out.json");

        DocumentWriter.Write(path, data);

        Assert.True(File.Exists(path));
        Assert.Contains("\"name\": \"X\"", File.ReadAllText(path));
    }

    [Fact]
    public void Write_UnknownExtension_Throws()
    {
        Assert.Throws<NotSupportedException>(() =>
            DocumentWriter.Write("file.unknown", new List<Product>()));
    }

    [Fact]
    public void WriteRead_Csv_Roundtrip()
    {
        var original = new List<Product>
        {
            new() { Name = "Widget", Price = 9.99m, Stock = 100 },
            new() { Name = "Gadget", Price = 24.50m, Stock = 50 },
        };
        var path = GetTempPath("roundtrip.csv");

        DocumentWriter.Write(path, original);
        var result = DocumentReader.Read<Product>(path);

        Assert.Equal(2, result.Count);
        Assert.Equal(original[0].Name, result[0].Name);
        Assert.Equal(original[0].Price, result[0].Price);
    }

    [Fact]
    public void WriteRead_Json_Roundtrip()
    {
        var original = new List<Product>
        {
            new() { Name = "A", Price = 1.0m, Stock = 1 },
            new() { Name = "B", Price = 2.0m, Stock = 2 },
        };
        var path = GetTempPath("roundtrip.json");

        DocumentWriter.Write(path, original);
        var result = DocumentReader.Read<Product>(path);

        Assert.Equal(2, result.Count);
        Assert.Equal(original[0].Name, result[0].Name);
    }

    [Fact]
    public void WriteRead_Excel_Roundtrip()
    {
        var original = new List<Product>
        {
            new() { Name = "Widget", Price = 9.99m, Stock = 100 },
            new() { Name = "Gadget", Price = 24.50m, Stock = 50 },
            new() { Name = "Doohickey", Price = 3.75m, Stock = 200 },
        };
        var path = GetTempPath("roundtrip.xlsx");

        DocumentWriter.Write(path, original);
        var result = DocumentReader.Read<Product>(path);

        Assert.Equal(3, result.Count);
        Assert.Equal("Widget", result[0].Name);
        Assert.Equal(9.99m, result[0].Price);
        Assert.Equal(100, result[0].Stock);
        Assert.Equal("Gadget", result[1].Name);
        Assert.Equal(24.50m, result[1].Price);
        Assert.Equal(50, result[1].Stock);
        Assert.Equal("Doohickey", result[2].Name);
        Assert.Equal(3.75m, result[2].Price);
        Assert.Equal(200, result[2].Stock);
    }

    [Fact]
    public void Write_Excel_EmptyList_Succeeds()
    {
        var path = GetTempPath("empty.xlsx");
        DocumentWriter.Write(path, new List<Product>());
        Assert.True(File.Exists(path));
    }

    [Fact]
    public void Write_Excel_SingleRow_Succeeds()
    {
        var data = new List<Product> { new() { Name = "Solo", Price = 1.0m, Stock = 1 } };
        var path = GetTempPath("single.xlsx");

        DocumentWriter.Write(path, data);
        var result = DocumentReader.Read<Product>(path);

        Assert.Single(result);
        Assert.Equal("Solo", result[0].Name);
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
