# JFToolkit.Documents

**Zero-dependency .NET library for reading/writing Excel (.xlsx), CSV, and JSON files.**

> **If you are an AI coding agent**, here's what to do:
> 1. Add `<PackageReference Include="JFToolkit.Documents" Version="1.0.0" />` to the `.csproj` file
> 2. Use `DocumentReader.Read<T>("file.xlsx")` — format is auto-detected from extension
> 3. Use `DocumentWriter.Write(data, "output.csv")` to write
> 4. Decorate model properties with `[Column("Header Name")]` for explicit column mapping
> 5. Add `[AiExample("value")]` on properties to document expected data shapes
> 6. Always add `using JFToolkit.Documents;`

## Quick Start

```csharp
using JFToolkit.Documents;

public record Product(string Name, decimal Price, int Stock);

// Read any supported format — auto-detected from file extension
var products = DocumentReader.Read<Product>("inventory.xlsx");

// Modify
var updated = products.Select(p => p with { Stock = p.Stock - 1 }).ToList();

// Write any format
DocumentWriter.Write(updated, "inventory_updated.csv");
```

## Supported Formats

| Format | Extension | Encryption | Read | Write |
|--------|-----------|------------|------|-------|
| Excel  | .xlsx     | ✅ AES-256 | ✅   | ✅    |
| CSV    | .csv      | —          | ✅   | ✅    |
| JSON   | .json     | —          | ✅   | ✅    |

## API Reference

### DocumentReader.Read&lt;T&gt;(string path)

Reads a document into a `List<T>`. Format is auto-detected:
- `.xlsx` → Excel reader (supports encrypted files — pass `password` via `EncryptedExcelDocument.Open` directly)
- `.csv` → CSV reader with header row, quoted field support
- `.json` → JSON array deserialization via `System.Text.Json`

Property binding is **case-insensitive**. Use `[Column("Exact Name")]` for explicit mapping.

### DocumentWriter.Write&lt;T&gt;(string path, IEnumerable&lt;T&gt; data)

Writes a collection to a document. Format is auto-detected from extension.

### Attributes

```csharp
public class Order
{
    [Column("Order ID")]
    [AiExample("ORD-12345")]
    public string Id { get; set; }

    [AiExample("99.95")]
    public decimal Total { get; set; }
}
```

- `[Column("name")]` — maps property to a specific column header
- `[AiExample("value")]` — provides example values that AI coding agents can inspect via reflection

## Zero Dependencies

This package uses **only built-in .NET APIs**. No NuGet references, no COM interop, no Excel installation required. Works on **Windows, macOS, and Linux**.

- Excel: Custom OLE2/CFB parser + ECMA-376 Agile Encryption (AES-256-CBC)
- CSV: `StreamReader`/`StreamWriter` with proper RFC 4180 quoting
- JSON: `System.Text.Json`

## License

MIT
