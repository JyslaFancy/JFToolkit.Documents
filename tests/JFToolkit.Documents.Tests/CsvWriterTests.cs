using JFToolkit.Documents;
using Xunit;

namespace JFToolkit.Documents.Tests;

public class CsvWriterTests
{
    [Fact]
    public void Write_BasicData_CreatesValidCsv()
    {
        var data = new List<Product>
        {
            new() { Name = "Widget", Price = 9.99m, Stock = 100 },
            new() { Name = "Gadget", Price = 24.50m, Stock = 50 },
        };
        var path = GetTempPath("write_basic.csv");

        Csv.CsvWriter.Write(path, data);

        var lines = File.ReadAllLines(path);
        Assert.Equal(3, lines.Length); // header + 2 data
        Assert.Equal("Name,Price,Stock", lines[0]);
        Assert.Equal("Widget,9.99,100", lines[1]);
        Assert.Equal("Gadget,24.50,50", lines[2]);
    }

    [Fact]
    public void Write_EmptyList_CreatesHeaderOnly()
    {
        var path = GetTempPath("write_empty.csv");
        Csv.CsvWriter.Write(path, new List<Product>());

        var lines = File.ReadAllLines(path);
        Assert.Single(lines);
        Assert.Equal("Name,Price,Stock", lines[0]);
    }

    [Fact]
    public void Write_NoHeader_SkipsHeaderRow()
    {
        var data = new List<Product> { new() { Name = "X", Price = 1, Stock = 1 } };
        var path = GetTempPath("write_noheader.csv");

        Csv.CsvWriter.Write(path, data, new Csv.CsvOptions { HasHeaderRow = false });

        var lines = File.ReadAllLines(path);
        Assert.Single(lines);
        Assert.Equal("X,1,1", lines[0]);
    }

    [Fact]
    public void Write_SemicolonDelimiter_UsesDelimiter()
    {
        var data = new List<Product> { new() { Name = "X", Price = 1, Stock = 1 } };
        var path = GetTempPath("write_semicolon.csv");

        Csv.CsvWriter.Write(path, data, new Csv.CsvOptions { Delimiter = ';' });

        var lines = File.ReadAllLines(path);
        Assert.Equal("Name;Price;Stock", lines[0]);
        Assert.Equal("X;1;1", lines[1]);
    }

    [Fact]
    public void Write_FieldsWithCommas_GetsQuoted()
    {
        var data = new List<QuotedModel>
        {
            new() { Name = "Hello", Description = "contains, comma" }
        };
        var path = GetTempPath("write_quoted.csv");

        Csv.CsvWriter.Write(path, data);

        var content = File.ReadAllText(path);
        Assert.Contains("\"contains, comma\"", content);
    }

    [Fact]
    public void Write_FieldsWithQuotes_GetsEscaped()
    {
        var data = new List<QuotedModel>
        {
            new() { Name = "Hello", Description = "has \"quote\" inside" }
        };
        var path = GetTempPath("write_escaped.csv");

        Csv.CsvWriter.Write(path, data);

        var content = File.ReadAllText(path);
        // Quotes are doubled for escaping
        Assert.Contains("\"has \"\"quote\"\" inside\"", content);
    }

    [Fact]
    public void Write_ColumnAttribute_UsesCustomHeader()
    {
        var data = new List<ColumnMappedModel>
        {
            new() { Name = "Widget", Price = 9.99m, Stock = 100 }
        };
        var path = GetTempPath("write_colattr.csv");

        Csv.CsvWriter.Write(path, data);

        var lines = File.ReadAllLines(path);
        Assert.Equal("Item Name,Item Price,Item Stock", lines[0]);
    }

    [Fact]
    public void Write_Roundtrip_PreservesData()
    {
        var original = new List<Product>
        {
            new() { Name = "A", Price = 1.23m, Stock = 10 },
            new() { Name = "B", Price = 4.56m, Stock = 20 },
            new() { Name = "C", Price = 7.89m, Stock = 30 },
        };
        var path = GetTempPath("write_roundtrip.csv");

        Csv.CsvWriter.Write(path, original);
        var roundtripped = Csv.CsvReader.Read<Product>(path);

        Assert.Equal(original.Count, roundtripped.Count);
        for (int i = 0; i < original.Count; i++)
        {
            Assert.Equal(original[i].Name, roundtripped[i].Name);
            Assert.Equal(original[i].Price, roundtripped[i].Price);
            Assert.Equal(original[i].Stock, roundtripped[i].Stock);
        }
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

    public class QuotedModel
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
    }

    public class ColumnMappedModel
    {
        [Column("Item Name")]
        public string Name { get; set; } = "";

        [Column("Item Price")]
        public decimal Price { get; set; }

        [Column("Item Stock")]
        public int Stock { get; set; }
    }
}
