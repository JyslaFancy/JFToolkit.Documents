using JFToolkit.Documents;
using Xunit;

namespace JFToolkit.Documents.Tests;

public class CsvReaderTests
{
    [Fact]
    public void Read_BasicCsv_ReturnsCorrectData()
    {
        var csv = "Name,Price,Stock\nWidget,9.99,100\nGadget,24.50,50";
        var path = GetTempPath("basic.csv");
        File.WriteAllText(path, csv);

        var result = Csv.CsvReader.Read<Product>(path);

        Assert.Equal(2, result.Count);
        Assert.Equal("Widget", result[0].Name);
        Assert.Equal(9.99m, result[0].Price);
        Assert.Equal(100, result[0].Stock);
        Assert.Equal("Gadget", result[1].Name);
        Assert.Equal(24.50m, result[1].Price);
        Assert.Equal(50, result[1].Stock);
    }

    [Fact]
    public void Read_EmptyFile_ReturnsEmptyList()
    {
        var path = GetTempPath("empty.csv");
        File.WriteAllText(path, "");

        var result = Csv.CsvReader.Read<Product>(path);

        Assert.Empty(result);
    }

    [Fact]
    public void Read_HeaderOnly_ReturnsEmptyList()
    {
        var path = GetTempPath("headeronly.csv");
        File.WriteAllText(path, "Name,Price,Stock");

        var result = Csv.CsvReader.Read<Product>(path);

        Assert.Empty(result);
    }

    [Fact]
    public void Read_QuotedFields_HandlesCorrectly()
    {
        var csv = "Name,Description\nWidget,\"Contains, comma\"\nGadget,\"Has \"\"quotes\"\" inside\"";
        var path = GetTempPath("quoted.csv");
        File.WriteAllText(path, csv);

        var result = Csv.CsvReader.Read<QuotedModel>(path);

        Assert.Equal(2, result.Count);
        Assert.Equal("Widget", result[0].Name);
        Assert.Equal("Contains, comma", result[0].Description);
        Assert.Equal("Gadget", result[1].Name);
        Assert.Equal("Has \"quotes\" inside", result[1].Description);
    }

    [Fact]
    public void Read_SemicolonDelimiter_Works()
    {
        var csv = "Name;Price;Stock\nWidget;9.99;100";
        var path = GetTempPath("semicolon.csv");
        File.WriteAllText(path, csv);

        var options = new Csv.CsvOptions { Delimiter = ';' };
        var result = Csv.CsvReader.Read<Product>(path, options);

        Assert.Single(result);
        Assert.Equal("Widget", result[0].Name);
        Assert.Equal(9.99m, result[0].Price);
    }

    [Fact]
    public void Read_CaseInsensitiveHeaders_MatchesProperties()
    {
        var csv = "NAME,price,STOCK\nWidget,9.99,100";
        var path = GetTempPath("case.csv");
        File.WriteAllText(path, csv);

        var result = Csv.CsvReader.Read<Product>(path);

        Assert.Single(result);
        Assert.Equal("Widget", result[0].Name);
        Assert.Equal(9.99m, result[0].Price);
        Assert.Equal(100, result[0].Stock);
    }

    [Fact]
    public void Read_ColumnAttribute_MapsCorrectly()
    {
        var csv = "Item Name,Item Price,Item Stock\nWidget,9.99,100";
        var path = GetTempPath("colattr.csv");
        File.WriteAllText(path, csv);

        var result = Csv.CsvReader.Read<ColumnMappedModel>(path);

        Assert.Single(result);
        Assert.Equal("Widget", result[0].Name);
        Assert.Equal(9.99m, result[0].Price);
        Assert.Equal(100, result[0].Stock);
    }

    [Fact]
    public void Read_MixedTypes_ParsesCorrectly()
    {
        var csv = "Text,Integer,Decimal,Bool,Date\nHello,42,3.14,true,2024-01-15";
        var path = GetTempPath("types.csv");
        File.WriteAllText(path, csv);

        var result = Csv.CsvReader.Read<MixedTypesModel>(path);

        Assert.Single(result);
        Assert.Equal("Hello", result[0].Text);
        Assert.Equal(42, result[0].Integer);
        Assert.Equal(3.14m, result[0].Decimal);
        Assert.True(result[0].Bool);
        Assert.Equal(new DateTime(2024, 1, 15), result[0].Date);
    }

    [Fact]
    public void Read_ExtraColumns_Ignored()
    {
        var csv = "Name,Price,Stock,Extra1,Extra2\nWidget,9.99,100,ignored,also_ignored";
        var path = GetTempPath("extra.csv");
        File.WriteAllText(path, csv);

        var result = Csv.CsvReader.Read<Product>(path);

        Assert.Single(result);
        Assert.Equal("Widget", result[0].Name);
    }

    [Fact]
    public void Read_MissingColumns_LeavesDefault()
    {
        var csv = "Name\nWidget";
        var path = GetTempPath("missing.csv");
        File.WriteAllText(path, csv);

        var result = Csv.CsvReader.Read<Product>(path);

        Assert.Single(result);
        Assert.Equal("Widget", result[0].Name);
        Assert.Equal(0m, result[0].Price);   // default
        Assert.Equal(0, result[0].Stock);     // default
    }

    [Fact]
    public void Read_EmptyLines_Skipped()
    {
        var csv = "Name,Price,Stock\n\nWidget,9.99,100\n\n\nGadget,24.50,50\n";
        var path = GetTempPath("emptylines.csv");
        File.WriteAllText(path, csv);

        var result = Csv.CsvReader.Read<Product>(path);

        Assert.Equal(2, result.Count);
    }

    private static string GetTempPath(string name)
    {
        var dir = Path.Combine(Path.GetTempPath(), "jftoolkit_tests");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, name);
    }

    // ── Test models ──

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

    public class MixedTypesModel
    {
        public string Text { get; set; } = "";
        public int Integer { get; set; }
        public decimal Decimal { get; set; }
        public bool Bool { get; set; }
        public DateTime Date { get; set; }
    }
}
