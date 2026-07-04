using JFToolkit.Documents;

var testDir = Path.Combine(Path.GetTempPath(), "jftoolkit_verify");
Directory.CreateDirectory(testDir);

int passed = 0, failed = 0;

void Check(bool cond, string msg)
{
    if (cond) { passed++; Console.WriteLine($"  ✅ {msg}"); }
    else { failed++; Console.WriteLine($"  ❌ {msg}"); }
}

Console.WriteLine("🔬 JFToolkit.Documents — verifiseringstest\n");

// ──── Data ────
var products = new List<Product>
{
    new() { Name = "Tastatur", Pris = 899.50m, Antall = 12 },
    new() { Name = "Mus", Pris = 349m, Antall = 45 },
    new() { Name = "Skjerm", Pris = 2499m, Antall = 3 },
};

// ──── CSV ────
Console.WriteLine("📄 CSV:");
var csvPath = Path.Combine(testDir, "test.csv");
DocumentWriter.Write(csvPath, products);
var csvRead = DocumentReader.Read<Product>(csvPath);
Check(csvRead.Count == 3, "Skriv og les 3 rader");
Check(csvRead[0].Name == "Tastatur", "Første rad — navn");
Check(csvRead[1].Pris == 349m, "Andre rad — pris");
Check(csvRead[2].Antall == 3, "Tredje rad — antall");

// ──── JSON ────
Console.WriteLine("📄 JSON:");
var jsonPath = Path.Combine(testDir, "test.json");
DocumentWriter.Write(jsonPath, products);
var jsonRead = DocumentReader.Read<Product>(jsonPath);
Check(jsonRead.Count == 3, "Skriv og les 3 rader");
Check(jsonRead[0].Name == "Tastatur", "Første rad — navn");

// ──── Excel ────
Console.WriteLine("📄 Excel:");
var xlsxPath = Path.Combine(testDir, "test.xlsx");
DocumentWriter.Write(xlsxPath, products);
var xlsxRead = DocumentReader.Read<Product>(xlsxPath);
Check(xlsxRead.Count == 3, "Skriv og les 3 rader");
Check(xlsxRead[0].Name == "Tastatur", "Første rad — navn");
Check(xlsxRead[2].Pris == 2499m, "Tredje rad — pris");

// ──── Vis Excel-fil ────
Console.WriteLine($"📂 Excel-fil: {xlsxPath}");

// ──── Cleanup ────
try { Directory.Delete(testDir, true); } catch { }

Console.WriteLine($"\n━━━ {passed}/{passed + failed} testar bestått ━━━");
if (failed > 0)
{
    Console.WriteLine("❌ Nokre testar feila!");
    Environment.Exit(1);
}
else
{
    Console.WriteLine("🎉 Alt fungerer!");
}

public class Product
{
    public string Name { get; set; } = "";
    public decimal Pris { get; set; }
    public int Antall { get; set; }
}
