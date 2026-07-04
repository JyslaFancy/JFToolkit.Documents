using JFToolkit.Documents;

// ── Test model ──
var sampleData = new List<Product>
{
    new() { Name = "Widget", Price = 9.99m, Stock = 100 },
    new() { Name = "Gadget", Price = 24.50m, Stock = 50 },
    new() { Name = "Doohickey", Price = 3.75m, Stock = 200 },
};

var testDir = Path.Combine(Path.GetTempPath(), "jftoolkit_demo");
Directory.CreateDirectory(testDir);

int passed = 0;
int failed = 0;

void Assert(bool condition, string test)
{
    if (condition) { passed++; Console.WriteLine($"  ✓ {test}"); }
    else { failed++; Console.WriteLine($"  ✗ FAIL: {test}"); }
}

try
{
    // ── CSV roundtrip ──
    Console.WriteLine("CSV roundtrip:");
    var csvPath = Path.Combine(testDir, "test.csv");
    DocumentWriter.Write(csvPath, sampleData);
    Assert(File.Exists(csvPath), "CSV file created");

    var csvRead = DocumentReader.Read<Product>(csvPath);
    Assert(csvRead.Count == 3, $"CSV read count: {csvRead.Count}");
    Assert(csvRead[0].Name == "Widget", $"CSV row 0 Name: {csvRead[0].Name}");
    Assert(csvRead[1].Price == 24.50m, $"CSV row 1 Price: {csvRead[1].Price}");
    Assert(csvRead[2].Stock == 200, $"CSV row 2 Stock: {csvRead[2].Stock}");

    // ── JSON roundtrip ──
    Console.WriteLine("JSON roundtrip:");
    var jsonPath = Path.Combine(testDir, "test.json");
    DocumentWriter.Write(jsonPath, sampleData);
    Assert(File.Exists(jsonPath), "JSON file created");

    var jsonRead = DocumentReader.Read<Product>(jsonPath);
    Assert(jsonRead.Count == 3, $"JSON read count: {jsonRead.Count}");
    Assert(jsonRead[0].Name == "Widget", $"JSON row 0 Name: {jsonRead[0].Name}");
    Assert(jsonRead[1].Price == 24.50m, $"JSON row 1 Price: {jsonRead[1].Price}");
    Assert(jsonRead[2].Stock == 200, $"JSON row 2 Stock: {jsonRead[2].Stock}");

    // ── Excel roundtrip ──
    Console.WriteLine("Excel roundtrip:");
    var xlsxPath = Path.Combine(testDir, "test.xlsx");
    DocumentWriter.Write(xlsxPath, sampleData);
    Assert(File.Exists(xlsxPath), "Excel file created");

    var xlsxRead = DocumentReader.Read<Product>(xlsxPath);
    Assert(xlsxRead.Count == 3, $"Excel read count: {xlsxRead.Count}");
    Assert(xlsxRead[0].Name == "Widget", $"Excel row 0 Name: {xlsxRead[0].Name}");
    Assert(xlsxRead[1].Price == 24.50m, $"Excel row 1 Price: {xlsxRead[1].Price}");
    Assert(xlsxRead[2].Stock == 200, $"Excel row 2 Stock: {xlsxRead[2].Stock}");

    // ── Edge cases ──
    Console.WriteLine("Edge cases:");
    
    // Empty data
    var emptyPath = Path.Combine(testDir, "empty.csv");
    DocumentWriter.Write(emptyPath, new List<Product>());
    var emptyRead = DocumentReader.Read<Product>(emptyPath);
    Assert(emptyRead.Count == 0, "Empty CSV roundtrip");

    // Single row
    var singlePath = Path.Combine(testDir, "single.csv");
    DocumentWriter.Write(singlePath, new List<Product> { new() { Name = "Solo", Price = 1.0m, Stock = 1 } });
    var singleRead = DocumentReader.Read<Product>(singlePath);
    Assert(singleRead.Count == 1, "Single row CSV");
    Assert(singleRead[0].Name == "Solo", "Single row Name");

    // File extension detection
    try
    {
        DocumentReader.Read<Product>("test.unknown");
        Assert(false, "Should have thrown for unknown format");
    }
    catch (NotSupportedException)
    {
        Assert(true, "Unknown format throws NotSupportedException");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"\nUNHANDLED ERROR: {ex}");
    failed++;
}

// ── Cleanup ──
try { Directory.Delete(testDir, recursive: true); } catch { }

Console.WriteLine($"\n─── Results: {passed} passed, {failed} failed ───");

if (failed > 0)
{
    Console.WriteLine("SOME TESTS FAILED!");
    Environment.Exit(1);
}
else
{
    Console.WriteLine("All tests passed!");
}

// ── Model ──
public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public int Stock { get; set; }
}
