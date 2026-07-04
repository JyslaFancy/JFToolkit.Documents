namespace JFToolkit.Documents.Excel;

/// <summary>
/// Opens encrypted or unencrypted Excel files for programmatic modification,
/// then saves with the same (or a new) password.
/// 
/// Usage:
///   using var doc = EncryptedExcelDocument.Open(@"C:\data\budget.xlsx", "secret123");
///   doc.SetCellValue("Sheet1", 3, 2, "New Value");
///   doc.SaveAs(@"C:\data\budget_updated.xlsx");
/// 
/// Zero dependencies. Pure C# — no COM interop, no Excel installation required.
/// Works on Windows, macOS, and Linux.
/// </summary>
public class EncryptedExcelDocument : IDisposable
{
    private readonly XlsxCellStore _store;
    private readonly string? _password;
    private string _filePath;
    private bool _isReadOnly;
    private bool _disposed;

    private EncryptedExcelDocument(XlsxCellStore store, string filePath,
        string? password, bool isReadOnly = false)
    {
        _store = store;
        _password = password;
        _filePath = filePath;
        _isReadOnly = isReadOnly;
    }

    // ── Public API ──

    /// <summary>
    /// Open an Excel file. Provide password for encrypted files, null for unencrypted.
    /// For encrypted .xlsx files, the password is used to decrypt the package.
    /// </summary>
    /// <param name="path">Full path to the .xlsx file.</param>
    /// <param name="password">File-open password, or null if the file is not encrypted.</param>
    /// <param name="readOnly">If true, the file cannot be saved back (Save will throw).</param>
    public static EncryptedExcelDocument Open(string path, string? password = null,
        bool readOnly = false)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Excel file not found: {path}");

        byte[] zipBytes;

        using (var fs = File.OpenRead(path))
        {
            if (AgileEncryption.IsOle2File(fs))
            {
                // Encrypted file — OLE2 container
                if (password == null)
                    throw new InvalidOperationException(
                        "This file is encrypted. Please provide a password.");

                fs.Position = 0;
                using var ole = OleStorage.Read(fs);
                var encInfo = ole.ReadStream("EncryptionInfo");
                var encryptedPackage = ole.ReadStream("EncryptedPackage");
                zipBytes = AgileEncryption.Decrypt(encInfo, encryptedPackage, password);
            }
            else
            {
                // Plain ZIP (unencrypted)
                fs.Position = 0;
                zipBytes = new byte[fs.Length];
                fs.Read(zipBytes, 0, zipBytes.Length);
            }
        }

        var store = XlsxCellStore.Load(zipBytes);
        return new EncryptedExcelDocument(store, path, password, readOnly);
    }

    /// <summary>
    /// Number of sheets in the workbook.
    /// </summary>
    public int SheetCount => _store.SheetCount;

    /// <summary>
    /// The full path of the opened workbook.
    /// </summary>
    public string FilePath => _filePath;

    /// <summary>
    /// Get all sheet names.
    /// </summary>
    public string[] GetSheetNames() => _store.GetSheetNames();

    /// <summary>
    /// Get the value of a cell (1-based row and column).
    /// </summary>
    public object? GetCellValue(string sheetName, int row, int column)
        => _store.GetCellValue(sheetName, row, column);

    /// <summary>
    /// Get the value of a cell on the first sheet (1-based row and column).
    /// </summary>
    public object? GetCellValue(int row, int column)
    {
        var firstName = _store.FirstSheetName
            ?? throw new InvalidOperationException("Workbook has no sheets.");
        return _store.GetCellValue(firstName, row, column);
    }

    /// <summary>
    /// Set the value of a cell (1-based row and column).
    /// Accepts string, double, int, DateTime, bool — passed through to Excel.
    /// </summary>
    public void SetCellValue(string sheetName, int row, int column, object value)
    {
        ThrowIfReadOnly();
        _store.SetCellValue(sheetName, row, column, value);
    }

    /// <summary>
    /// Set the value of a cell on the first sheet (1-based row and column).
    /// </summary>
    public void SetCellValue(int row, int column, object value)
    {
        ThrowIfReadOnly();
        var firstName = _store.FirstSheetName
            ?? throw new InvalidOperationException("Workbook has no sheets.");
        _store.SetCellValue(firstName, row, column, value);
    }

    /// <summary>
    /// Save the workbook to its current path (overwrites original file).
    /// Re-encrypts with the original password if the file was encrypted.
    /// </summary>
    public void Save()
    {
        SaveAs(_filePath, _password);
    }

    /// <summary>
    /// Save the workbook to a new file. If the original was encrypted,
    /// the new file is encrypted with the same password.
    /// Pass null to save unencrypted.
    /// </summary>
    public void SaveAs(string newPath)
    {
        SaveAs(newPath, _password);
    }

    /// <summary>
    /// Save the workbook to a new file with a new password.
    /// Pass null for password to save unencrypted.
    /// </summary>
    public void SaveAs(string newPath, string? newPassword)
    {
        ThrowIfDisposed();
        ThrowIfReadOnly();

        var innerZip = _store.Save();

        if (newPassword is not null)
        {
            // Encrypt and wrap in OLE2 container
            var (encInfo, encPackage) = AgileEncryption.Encrypt(innerZip, newPassword);

            using var ms = new MemoryStream();
            OleStorage.Write(ms,
                ("EncryptionInfo", encInfo),
                ("EncryptedPackage", encPackage));

            File.WriteAllBytes(newPath, ms.ToArray());
        }
        else
        {
            // Write plain ZIP
            File.WriteAllBytes(newPath, innerZip);
        }

        _filePath = newPath;
    }

    /// <summary>
    /// Returns true if the file is an encrypted Excel file that this library
    /// can decrypt (OLE2 compound document with ECMA-376 Agile Encryption).
    /// </summary>
    public static bool IsEncrypted(string path)
    {
        using var fs = File.OpenRead(path);
        return AgileEncryption.IsOle2File(fs);
    }

    /// <summary>
    /// Returns true if the file is a valid Excel file (encrypted or not).
    /// </summary>
    public static bool IsExcelFile(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            var magic = new byte[8];
            fs.Read(magic, 0, 8);
            fs.Position = 0;

            // PK\x03\x04 = ZIP (unencrypted .xlsx)
            if (magic[0] == 0x50 && magic[1] == 0x4B)
                return true;

            // D0 CF 11 E0 = OLE2 (encrypted .xlsx)
            return AgileEncryption.IsOle2File(fs);
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(EncryptedExcelDocument));
    }

    private void ThrowIfReadOnly()
    {
        ThrowIfDisposed();
        if (_isReadOnly) throw new InvalidOperationException(
            "This document was opened in read-only mode and cannot be modified.");
    }
}
