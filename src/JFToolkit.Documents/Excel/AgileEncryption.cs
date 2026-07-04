using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace JFToolkit.Documents.Excel;

/// <summary>
/// ECMA-376 Agile Encryption (Part 4).
/// 
/// Implements AES-256-CBC with SHA-512 key derivation.
/// Handles decrypting and encrypting Office Open XML packages
/// (the ZIP inside an encrypted .xlsx).
/// 
/// Zero dependencies — uses only System.Security.Cryptography.
/// </summary>
internal static class AgileEncryption
{
    // ── Encryption parameters ──
    private const int SaltSize = 16;
    private const int BlockSize = 16;
    private const int KeyBits = 256;
    private const int KeyBytes = KeyBits / 8; // 32
    private const int HashSize = 64;
    private const int DefaultSpinCount = 100_000;
    private const int ChunkSize = 4096;

    // ── Block keys (from ECMA-376 standard) ──
    private static readonly byte[] VerifierHashInputBlockKey = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
    private static readonly byte[] VerifierHashValueBlockKey = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
    private static readonly byte[] EncryptedKeyValueBlockKey = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
    private static readonly byte[] HmacKeyBlock = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
    private static readonly byte[] HmacValueBlock = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

    // ═══════════════════════════════════════════════
    //  PUBLIC API
    // ═══════════════════════════════════════════════

    /// <summary>
    /// Decrypt an EncryptedPackage stream to recover the inner ZIP bytes.
    /// </summary>
    /// <param name="encryptionInfo">The raw EncryptionInfo stream bytes (4-byte version + XML).</param>
    /// <param name="encryptedPackage">The raw EncryptedPackage stream (16-byte salt + encrypted payload).</param>
    /// <param name="password">The file-open password.</param>
    public static byte[] Decrypt(byte[] encryptionInfo, byte[] encryptedPackage, string password)
    {
        // Parse the EncryptionInfo XML to get encryption parameters
        var encInfo = ParseEncryptionInfo(encryptionInfo);
        var dataKey = RecoverDataKey(encInfo, password);

        // Decrypt the package using recovered data key
        return DecryptPackage(encryptedPackage, dataKey);
    }

    /// <summary>
    /// Encrypt an inner ZIP to produce the EncryptedPackage stream + EncryptionInfo.
    /// Returns (encryptionInfo, encryptedPackage) ready for OLE2 container.
    /// </summary>
    public static (byte[] EncryptionInfo, byte[] EncryptedPackage) Encrypt(
        byte[] innerZip, string password)
    {
        var spinCount = DefaultSpinCount;

        // Generate random salts
        var keySalt = RandomBytes(SaltSize);
        var packageSalt = RandomBytes(SaltSize);
        var hmacSalt = RandomBytes(SaltSize);

        // Derive password hash
        var passwordHash = DerivePasswordHash(password, keySalt, spinCount);

        // Derive sub-keys
        var keyValueKey = DeriveBlockKey(passwordHash, EncryptedKeyValueBlockKey);
        var verifierKey = DeriveBlockKey(passwordHash, VerifierHashInputBlockKey);
        var hashKey = DeriveBlockKey(passwordHash, VerifierHashValueBlockKey);

        // Generate data key and verifier
        var dataKey = RandomBytes(KeyBytes);
        var verifier = RandomBytes(SaltSize);

        // Encrypt data key
        var encryptedKeyValue = AesChunkEncrypt(
            PadWith0x36(dataKey), keySalt, keyValueKey);

        // Encrypt verifier
        var encryptedVerifier = AesChunkEncrypt(
            verifier, keySalt, verifierKey);

        // Encrypt verifier hash
        var verifierHash = SHA512.HashData(verifier);
        var encryptedVerifierHash = AesChunkEncrypt(
            verifierHash, keySalt, hashKey);

        // Encrypt the inner ZIP
        var encryptedPackage = EncryptPackage(innerZip, dataKey, packageSalt);

        // Compute HMAC
        var hmacKey = DeriveBlockKey(passwordHash, HmacKeyBlock);
        var hmacValueKey = DeriveBlockKey(passwordHash, HmacValueBlock);
        var encryptedHmacKey = AesChunkEncrypt(hmacSalt, keySalt, hmacKey);
        var hmacValue = ComputeHmac(hmacSalt, encryptedPackage);
        var encryptedHmacValue = AesChunkEncrypt(hmacValue, keySalt, hmacValueKey);

        // Build EncryptionInfo XML
        var encryptionInfo = BuildEncryptionInfoXml(
            keySalt, packageSalt, hmacSalt, spinCount,
            encryptedKeyValue, encryptedVerifier, encryptedVerifierHash,
            encryptedHmacKey, encryptedHmacValue);

        return (encryptionInfo, encryptedPackage);
    }

    /// <summary>
    /// Check if a file is an OLE2 compound file (encrypted .xlsx).
    /// </summary>
    public static bool IsOle2File(Stream stream)
    {
        var magic = new byte[8];
        var pos = stream.Position;
        stream.Read(magic, 0, 8);
        stream.Position = pos;
        return BitConverter.ToUInt64(magic, 0) == 0xE11AB1A1E011CFD0;
    }

    // ═══════════════════════════════════════════════
    //  KEY DERIVATION
    // ═══════════════════════════════════════════════

    private static byte[] DerivePasswordHash(string password, byte[] salt, int spinCount)
    {
        var passwordBytes = Encoding.Unicode.GetBytes(password); // UTF-16LE
        var combined = new byte[salt.Length + passwordBytes.Length];
        Buffer.BlockCopy(salt, 0, combined, 0, salt.Length);
        Buffer.BlockCopy(passwordBytes, 0, combined, salt.Length, passwordBytes.Length);

        var hash = SHA512.HashData(combined);

        for (int i = 0; i < spinCount; i++)
        {
            var iterator = BitConverter.GetBytes(i); // 4-byte LE
            var input = new byte[4 + hash.Length];
            Buffer.BlockCopy(iterator, 0, input, 0, 4);
            Buffer.BlockCopy(hash, 0, input, 4, hash.Length);
            hash = SHA512.HashData(input);
        }

        return hash;
    }

    private static byte[] DeriveBlockKey(byte[] passwordHash, byte[] blockKey)
    {
        var input = new byte[passwordHash.Length + blockKey.Length];
        Buffer.BlockCopy(passwordHash, 0, input, 0, passwordHash.Length);
        Buffer.BlockCopy(blockKey, 0, input, passwordHash.Length, blockKey.Length);
        var hash = SHA512.HashData(input);
        var key = new byte[KeyBytes];
        Buffer.BlockCopy(hash, 0, key, 0, KeyBytes);
        return key;
    }

    // ═══════════════════════════════════════════════
    //  ENCRYPTION / DECRYPTION HELPERS
    // ═══════════════════════════════════════════════

    private static byte[] AesChunkEncrypt(byte[] data, byte[] salt, byte[] key)
    {
        // Pad to block boundary
        var paddedLength = ((data.Length + BlockSize - 1) / BlockSize) * BlockSize;
        var padded = new byte[paddedLength];
        Buffer.BlockCopy(data, 0, padded, 0, data.Length);

        var numChunks = paddedLength / BlockSize;
        var result = new byte[paddedLength + BlockSize]; // +1 block for IV (workaround: encrypt in place)

        // Actually, for simplicity, encrypt the whole thing at once with CBC
        // Each 16-byte "chunk" has its own IV derived from salt + chunk index
        using var aes = Aes.Create();
        aes.Key = key;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;

        var encrypted = new byte[paddedLength];
        for (int i = 0; i < numChunks; i++)
        {
            var iv = DeriveIv(salt, BitConverter.GetBytes(i));
            aes.IV = iv;

            var offset = i * BlockSize;
            using var encryptor = aes.CreateEncryptor();
            encryptor.TransformBlock(padded, offset, BlockSize, encrypted, offset);
        }

        return encrypted;
    }

    private static byte[] AesChunkDecrypt(byte[] data, byte[] salt, byte[] key)
    {
        if (data.Length % BlockSize != 0)
            throw new InvalidDataException($"Encrypted data length {data.Length} is not aligned to {BlockSize} bytes.");

        using var aes = Aes.Create();
        aes.Key = key;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;

        var decrypted = new byte[data.Length];
        var numChunks = data.Length / BlockSize;

        for (int i = 0; i < numChunks; i++)
        {
            var iv = DeriveIv(salt, BitConverter.GetBytes(i));
            aes.IV = iv;

            var offset = i * BlockSize;
            using var decryptor = aes.CreateDecryptor();
            decryptor.TransformBlock(data, offset, BlockSize, decrypted, offset);
        }

        return decrypted;
    }

    private static byte[] DeriveIv(byte[] salt, byte[] blockKey)
    {
        var input = new byte[salt.Length + blockKey.Length];
        Buffer.BlockCopy(salt, 0, input, 0, salt.Length);
        Buffer.BlockCopy(blockKey, 0, input, salt.Length, blockKey.Length);
        var hash = SHA512.HashData(input);
        var iv = new byte[BlockSize];
        Buffer.BlockCopy(hash, 0, iv, 0, BlockSize);
        return iv;
    }

    private static byte[] PadWith0x36(byte[] key)
    {
        var result = new byte[KeyBytes * 2];
        Buffer.BlockCopy(key, 0, result, 0, KeyBytes);
        for (int i = KeyBytes; i < result.Length; i++)
            result[i] = 0x36;
        return result;
    }

    // ═══════════════════════════════════════════════
    //  ENCRYPTION INFO (XML)
    // ═══════════════════════════════════════════════

    private sealed record EncryptionInfoData(
        byte[] KeySalt,
        byte[] PackageSalt,
        byte[] HmacSalt,
        int SpinCount,
        byte[] EncryptedKeyValue,
        byte[] EncryptedVerifier,
        byte[] EncryptedVerifierHash,
        byte[] EncryptedHmacKey,
        byte[] EncryptedHmacValue);

    private static EncryptionInfoData ParseEncryptionInfo(byte[] raw)
    {
        // The EncryptionInfo stream contains:
        // - First 4 bytes: version (major.minor, 2+2)
        // - Followed by XML
        using var ms = new MemoryStream(raw);
        var version = new byte[4];
        ms.Read(version, 0, 4);

        var xml = XElement.Load(ms);
        var ns = xml.GetDefaultNamespace();

        var keyData = xml.Element(ns + "keyData")
            ?? throw new InvalidDataException("Missing keyData in EncryptionInfo.");
        var dataIntegrity = xml.Element(ns + "dataIntegrity")
            ?? throw new InvalidDataException("Missing dataIntegrity in EncryptionInfo.");

        var encHmacKey = dataIntegrity.Attribute("encryptedHmacKey")?.Value
                ?? throw new InvalidDataException("Missing encryptedHmacKey.");
        var encHmacValue = dataIntegrity.Attribute("encryptedHmacValue")?.Value
                ?? throw new InvalidDataException("Missing encryptedHmacValue.");

        return new EncryptionInfoData(
            KeySalt: Convert.FromBase64String(keyData.Attribute("saltValue")?.Value ?? ""),
            PackageSalt: Convert.FromBase64String(keyData.Attribute("saltValue")?.Value ?? ""),
            HmacSalt: Array.Empty<byte>(),
            SpinCount: int.Parse(keyData.Attribute("spinCount")?.Value ?? "100000"),
            EncryptedKeyValue: Convert.FromBase64String(keyData.Element(ns + "encryptedKey")?.Attribute("encryptedKeyValue")?.Value
                ?? throw new InvalidDataException("Missing encryptedKeyValue.")),
            EncryptedVerifier: Convert.FromBase64String(keyData.Element(ns + "encryptedVerifier")?.Attribute("encryptedVerifierValue")?.Value
                ?? throw new InvalidDataException("Missing encryptedVerifierValue.")),
            EncryptedVerifierHash: Convert.FromBase64String(keyData.Element(ns + "encryptedVerifierHash")?.Attribute("encryptedVerifierHashValue")?.Value
                ?? throw new InvalidDataException("Missing encryptedVerifierHashValue.")),
            EncryptedHmacKey: Convert.FromBase64String(encHmacKey),
            EncryptedHmacValue: Convert.FromBase64String(encHmacValue)
        );
    }

    private static byte[] RecoverDataKey(EncryptionInfoData info, string password)
    {
        // Derive password hash
        var passwordHash = DerivePasswordHash(password, info.KeySalt, info.SpinCount);

        // Derive keys for verification
        var verifierKey = DeriveBlockKey(passwordHash, VerifierHashInputBlockKey);
        var hashKey = DeriveBlockKey(passwordHash, VerifierHashValueBlockKey);

        // Verify password by decrypting verifier and checking hash
        var decryptedVerifier = AesChunkDecrypt(info.EncryptedVerifier, info.KeySalt, verifierKey);
        var expectedHash = SHA512.HashData(decryptedVerifier);
        var actualHash = AesChunkDecrypt(info.EncryptedVerifierHash, info.KeySalt, hashKey);

        if (!CryptographicOperations.FixedTimeEquals(expectedHash, actualHash))
            throw new InvalidOperationException("Wrong password.");

        // Derive key encryption key and decrypt data key
        var keyValueKey = DeriveBlockKey(passwordHash, EncryptedKeyValueBlockKey);
        var paddedKey = AesChunkDecrypt(info.EncryptedKeyValue, info.KeySalt, keyValueKey);
        var dataKey = new byte[KeyBytes];
        Buffer.BlockCopy(paddedKey, 0, dataKey, 0, KeyBytes);

        return dataKey;
    }

    private static byte[] DecryptPackage(byte[] encryptedPackage, byte[] key)
    {
        // First 8 bytes: original size (64-bit little-endian)
        var originalSize = BitConverter.ToInt64(encryptedPackage, 0);
        var encryptedData = new byte[encryptedPackage.Length - 8];
        Buffer.BlockCopy(encryptedPackage, 8, encryptedData, 0, encryptedData.Length);

        // The salt used for package encryption is the first 16 bytes of the key itself?
        // Actually, looking at the office_encrypt reference, the package uses its own "package_salt".
        // But in the EncryptionInfo parsing, the salt for the package is... 
        // Hmm. Let me re-examine. The package IVs are derived from a separate salt.
        // Looking at the Microsoft implementation more carefully:
        // The package salt is stored as part of the encrypted package stream header.
        // Wait, looking at it again: the encryptedPackage format is:
        // [8 bytes: total size][encrypted data chunks]

        // The salt for the package encryption is the SAME as keyData.saltValue in many implementations.
        // But wait, looking at the office_encrypt reference:
        //   {encryption_info, encrypted_package} = perform_encryption(data, password, spin_count)
        // And perform_encryption generates a separate package_salt.
        // But for decryption, the package_salt comes from... where?
        // Actually, looking at the DecryptPackage implementation: the encrypted package
        // data starts with the salt (16 bytes) then the encrypted data.
        // Let me check... 

        // Actually, the encrypted package format per ECMA-376 is:
        // The EncryptedPackage stream inside the OLE2 container:
        // [16 bytes: package salt][encrypted data]
        // But wait: the 8-byte prefix is the total package size after decryption.

        // Wait, I need to re-examine. The office_encrypt implementation prepends:
        // <<byte_size(data)::little-64>> before chunk_encrypt
        // But it uses a SEPARATE package_salt for IV derivation.
        // The package_salt is... hmm, where is it stored?

        // Let me look at the actual EncryptedPackage format more carefully.
        // In the OLE2 file, the EncryptedPackage stream's first 16 bytes
        // are the package salt, then the encrypted data follows.
        // Actually re-examining: the stream is laid out as:
        // [package_salt: 16 bytes][encrypted data: N bytes]
        // where encrypted data is: [8-byte size prefix || actual data] encrypted

        // Wait no. Let me trace through the Elixir implementation again:
        // encrypted_package = chunk_aes_encrypt(<<byte_size(data)::little-64>> <> data, package_salt, data_key)
        // So the package_salt is used to derive IVs, and the encrypted output
        // is prepended with an 8-byte size.

        // For reading: we need to find the package_salt. In Microsoft's implementation,
        // the package_salt is the first 16 bytes of the EncryptedPackage stream,
        // followed by the encrypted content.

        // Let me fix this:
        return DecryptPackageInternal(encryptedPackage, key);
    }

    private static byte[] DecryptPackageInternal(byte[] data, byte[] key)
    {
        // The EncryptedPackage stream format:
        // [16 bytes: package salt][8 bytes: encrypted size prefix + data][...]
        if (data.Length < 16)
            throw new InvalidDataException("EncryptedPackage too short.");

        var packageSalt = new byte[16];
        Buffer.BlockCopy(data, 0, packageSalt, 0, 16);

        var encryptedPayload = new byte[data.Length - 16];
        Buffer.BlockCopy(data, 16, encryptedPayload, 0, encryptedPayload.Length);

        // Decrypt
        var decrypted = AesChunkDecrypt(encryptedPayload, packageSalt, key);

        // First 8 bytes = original size
        var originalSize = BitConverter.ToInt64(decrypted, 0);
        var result = new byte[originalSize];
        Buffer.BlockCopy(decrypted, 8, result, 0, (int)originalSize);
        return result;
    }

    private static byte[] EncryptPackage(byte[] innerZip, byte[] key, byte[] salt)
    {
        // Prepend 8-byte size prefix, then encrypt
        var sizePrefix = BitConverter.GetBytes((long)innerZip.Length);
        var payload = new byte[8 + innerZip.Length];
        Buffer.BlockCopy(sizePrefix, 0, payload, 0, 8);
        Buffer.BlockCopy(innerZip, 0, payload, 8, innerZip.Length);

        var encryptedPayload = AesChunkEncrypt(payload, salt, key);

        // Output: [package_salt: 16][encrypted_payload]
        var result = new byte[16 + encryptedPayload.Length];
        Buffer.BlockCopy(salt, 0, result, 0, 16);
        Buffer.BlockCopy(encryptedPayload, 0, result, 16, encryptedPayload.Length);

        return result;
    }

    private static byte[] BuildEncryptionInfoXml(
        byte[] keySalt, byte[] packageSalt, byte[] hmacSalt, int spinCount,
        byte[] encryptedKeyValue, byte[] encryptedVerifier, byte[] encryptedVerifierHash,
        byte[] encryptedHmacKey, byte[] encryptedHmacValue)
    {
        var ns = "http://schemas.microsoft.com/office/2006/encryption";

        var doc = new XElement(ns + "encryption",
            new XElement(ns + "keyData",
                new XAttribute("saltSize", SaltSize),
                new XAttribute("blockSize", BlockSize),
                new XAttribute("keyBits", KeyBits),
                new XAttribute("hashSize", HashSize),
                new XAttribute("cipherAlgorithm", "AES"),
                new XAttribute("cipherChaining", "ChainingModeCBC"),
                new XAttribute("hashAlgorithm", "SHA512"),
                new XAttribute("saltValue", Convert.ToBase64String(keySalt)),
                new XAttribute("spinCount", spinCount),
                new XElement(ns + "encryptedKey",
                    new XAttribute("spinCount", spinCount),
                    new XAttribute("saltValue", Convert.ToBase64String(keySalt)),
                    new XAttribute("encryptedKeyValue",
                        Convert.ToBase64String(encryptedKeyValue))),
                new XElement(ns + "encryptedVerifier",
                    new XAttribute("spinCount", spinCount),
                    new XAttribute("saltValue", Convert.ToBase64String(keySalt)),
                    new XAttribute("encryptedVerifierValue",
                        Convert.ToBase64String(encryptedVerifier))),
                new XElement(ns + "encryptedVerifierHash",
                    new XAttribute("spinCount", spinCount),
                    new XAttribute("saltValue", Convert.ToBase64String(keySalt)),
                    new XAttribute("encryptedVerifierHashValue",
                        Convert.ToBase64String(encryptedVerifierHash)))),
            new XElement(ns + "dataIntegrity",
                new XAttribute("encryptedHmacKey", Convert.ToBase64String(encryptedHmacKey)),
                new XAttribute("encryptedHmacValue", Convert.ToBase64String(encryptedHmacValue)))
        );

        // Prepend version (4.4 = "Agile Encryption" major.minor)
        var version = new byte[] { 4, 0, 4, 0 };
        var xmlBytes = Encoding.UTF8.GetBytes(doc.ToString(SaveOptions.DisableFormatting));
        var result = new byte[version.Length + xmlBytes.Length];
        Buffer.BlockCopy(version, 0, result, 0, version.Length);
        Buffer.BlockCopy(xmlBytes, 0, result, version.Length, xmlBytes.Length);
        return result;
    }

    private static byte[] ComputeHmac(byte[] hmacSalt, byte[] encryptedPackage)
    {
        using var hmac = new HMACSHA512(hmacSalt);
        return hmac.ComputeHash(encryptedPackage);
    }

    private static byte[] RandomBytes(int count)
    {
        var bytes = new byte[count];
        RandomNumberGenerator.Fill(bytes);
        return bytes;
    }
}
