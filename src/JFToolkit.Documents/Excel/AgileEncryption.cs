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
    public static byte[] Decrypt(byte[] encryptionInfo, byte[] encryptedPackage, string password)
    {
        var encInfo = ParseEncryptionInfo(encryptionInfo);
        var dataKey = RecoverDataKey(encInfo, password);
        return DecryptPackage(encryptedPackage, dataKey, encInfo.KeyDataSalt);
    }

    /// <summary>
    /// Encrypt an inner ZIP to produce the EncryptedPackage stream + EncryptionInfo.
    /// Returns (encryptionInfo, encryptedPackage) ready for OLE2 container.
    /// </summary>
    public static (byte[] EncryptionInfo, byte[] EncryptedPackage) Encrypt(
        byte[] innerZip, string password)
    {
        // Salt for password-based key derivation
        var passwordSalt = RandomBytes(SaltSize);
        // Separate salt for package encryption (stored in keyData)
        var keyDataSalt = RandomBytes(SaltSize);

        // Derive password hash
        var passwordHash = DerivePasswordHash(password, passwordSalt, DefaultSpinCount);

        // Derive sub-keys
        var verifierKey = DeriveBlockKey(passwordHash, VerifierHashInputBlockKey);
        var hashKey = DeriveBlockKey(passwordHash, VerifierHashValueBlockKey);
        var keyValueKey = DeriveBlockKey(passwordHash, EncryptedKeyValueBlockKey);

        // Generate random verifier (16 bytes), pad to block, encrypt
        var verifier = RandomBytes(SaltSize);
        var paddedVerifier = PadToBlockSize(verifier, BlockSize);
        var encryptedVerifier = AesCbcEncrypt(paddedVerifier, verifierKey, passwordSalt);

        // Hash the ORIGINAL verifier (not padded), pad, encrypt
        var verifierHash = SHA512.HashData(verifier);
        var paddedHash = PadToBlockSize(verifierHash, BlockSize);
        var encryptedVerifierHash = AesCbcEncrypt(paddedHash, hashKey, passwordSalt);

        // Generate random data key (16 bytes of entropy), pad with 0x36 to KeyBytes
        var dataKeyEntropy = RandomBytes(SaltSize);
        var paddedKey = PadWith0x36(dataKeyEntropy);
        var encryptedKeyValue = AesCbcEncrypt(paddedKey, keyValueKey, passwordSalt);

        // Encrypt the inner ZIP using keyDataSalt for IV derivation
        var encryptedPackage = EncryptPackage(innerZip, paddedKey, keyDataSalt);

        // Compute HMAC for integrity (matching msoffcrypto implementation)
        var hmacSalt = RandomBytes(HashSize); // 64 bytes of random
        var hmacIv1 = DeriveIv(keyDataSalt, HmacKeyBlock); // IV from keyDataSalt + blockKey
        var hmacIv2 = DeriveIv(keyDataSalt, HmacValueBlock);
        var encryptedHmacKey = AesCbcEncrypt(PadToBlockSize(hmacSalt, BlockSize), paddedKey, hmacIv1);
        var hmacValue = ComputeHmac(hmacSalt, encryptedPackage);
        var encryptedHmacValue = AesCbcEncrypt(PadToBlockSize(hmacValue, BlockSize), paddedKey, hmacIv2);

        // Build EncryptionInfo XML
        var encryptionInfo = BuildEncryptionInfoXml(
            keyDataSalt, passwordSalt, DefaultSpinCount,
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
        stream.ReadExactly(magic, 0, 8);
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
            var iterator = BitConverter.GetBytes(i);
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

    private static byte[] AesCbcEncrypt(byte[] data, byte[] key, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.Key = key;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        aes.IV = iv;

        using var encryptor = aes.CreateEncryptor();
        return encryptor.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] AesCbcDecrypt(byte[] data, byte[] key, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.Key = key;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        aes.IV = iv;

        using var decryptor = aes.CreateDecryptor();
        return decryptor.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] DeriveIv(byte[] salt, int chunkIndex)
    {
        var blockKey = BitConverter.GetBytes(chunkIndex);
        return DeriveIv(salt, blockKey);
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

    private static byte[] PadToBlockSize(byte[] data, int blockSize)
    {
        var paddedLength = ((data.Length + blockSize - 1) / blockSize) * blockSize;
        if (paddedLength == data.Length) return data;
        var result = new byte[paddedLength];
        Buffer.BlockCopy(data, 0, result, 0, data.Length);
        return result;
    }

    private static byte[] PadWith0x36(byte[] key)
    {
        // Pad to KeyBytes (32 for AES-256) with 0x36, per ECMA-376 / msoffcrypto
        var result = new byte[KeyBytes];
        Buffer.BlockCopy(key, 0, result, 0, key.Length);
        for (int i = key.Length; i < result.Length; i++)
            result[i] = 0x36;
        return result;
    }

    // ═══════════════════════════════════════════════
    //  ENCRYPTION INFO (XML)
    // ═══════════════════════════════════════════════

    private sealed record EncryptionInfoData(
        byte[] KeyDataSalt,
        byte[] PasswordSalt,
        int SpinCount,
        byte[] EncryptedKeyValue,
        byte[] EncryptedVerifierHashInput,
        byte[] EncryptedVerifierHashValue,
        byte[] EncryptedHmacKey,
        byte[] EncryptedHmacValue);

    private static EncryptionInfoData ParseEncryptionInfo(byte[] raw)
    {
        // First 8 bytes: version (2+2) + flags (4), then XML
        using var ms = new MemoryStream(raw);
        var header = new byte[8];
        ms.ReadExactly(header, 0, 8);

        var xml = XElement.Load(ms);
        var ns = xml.GetDefaultNamespace();
        XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        // keyData is self-closing with all attributes
        var keyData = xml.Element(ns + "keyData")
            ?? throw new InvalidDataException("Missing keyData in EncryptionInfo.");
        var dataIntegrity = xml.Element(ns + "dataIntegrity")
            ?? throw new InvalidDataException("Missing dataIntegrity in EncryptionInfo.");

        // keyEncryptors is a sibling of keyData, inside a keyEncryptor
        var keyEncryptors = xml.Element(ns + "keyEncryptors");
        var keyEncryptor = keyEncryptors?.Elements(ns + "keyEncryptor")
            .FirstOrDefault(ke => ke.Attribute("uri")?.Value == p.NamespaceName)
            ?? throw new InvalidDataException("Missing p:keyEncryptor in EncryptionInfo.");

        // All crypto values are attributes on a single p:encryptedKey element
        var encKey = keyEncryptor.Element(p + "encryptedKey")
            ?? throw new InvalidDataException("Missing p:encryptedKey.");

        return new EncryptionInfoData(
            KeyDataSalt: Convert.FromBase64String(keyData.Attribute("saltValue")?.Value ?? ""),
            PasswordSalt: Convert.FromBase64String(encKey.Attribute("saltValue")?.Value ?? ""),
            SpinCount: int.Parse(encKey.Attribute("spinCount")?.Value ?? "100000"),
            EncryptedKeyValue: Convert.FromBase64String(encKey.Attribute("encryptedKeyValue")?.Value
                ?? throw new InvalidDataException("Missing encryptedKeyValue.")),
            EncryptedVerifierHashInput: Convert.FromBase64String(encKey.Attribute("encryptedVerifierHashInput")?.Value
                ?? throw new InvalidDataException("Missing encryptedVerifierHashInput.")),
            EncryptedVerifierHashValue: Convert.FromBase64String(encKey.Attribute("encryptedVerifierHashValue")?.Value
                ?? throw new InvalidDataException("Missing encryptedVerifierHashValue.")),
            EncryptedHmacKey: Convert.FromBase64String(dataIntegrity.Attribute("encryptedHmacKey")?.Value
                ?? throw new InvalidDataException("Missing encryptedHmacKey.")),
            EncryptedHmacValue: Convert.FromBase64String(dataIntegrity.Attribute("encryptedHmacValue")?.Value
                ?? throw new InvalidDataException("Missing encryptedHmacValue."))
        );
    }

    private static byte[] RecoverDataKey(EncryptionInfoData info, string password)
    {
        try
        {
            // Derive password hash using the password node's salt
            var passwordHash = DerivePasswordHash(password, info.PasswordSalt, info.SpinCount);

            // Derive keys
            var verifierKey = DeriveBlockKey(passwordHash, VerifierHashInputBlockKey);
            var hashKey = DeriveBlockKey(passwordHash, VerifierHashValueBlockKey);

            // Verify password: decrypt verifier, hash it, compare with decrypted hash
            var decryptedVerifier = AesCbcDecrypt(info.EncryptedVerifierHashInput, verifierKey, info.PasswordSalt);
            var expectedHash = SHA512.HashData(decryptedVerifier);
            var actualHash = AesCbcDecrypt(info.EncryptedVerifierHashValue, hashKey, info.PasswordSalt);

            // Truncate to original hash size
            if (!CryptographicOperations.FixedTimeEquals(
                expectedHash.AsSpan(0, HashSize),
                actualHash.AsSpan(0, HashSize)))
                throw new InvalidOperationException("Wrong password.");

            // Derive key encryption key and recover data key
            var keyValueKey = DeriveBlockKey(passwordHash, EncryptedKeyValueBlockKey);
            var paddedKey = AesCbcDecrypt(info.EncryptedKeyValue, keyValueKey, info.PasswordSalt);
            var dataKey = new byte[KeyBytes];
            Buffer.BlockCopy(paddedKey, 0, dataKey, 0, KeyBytes);

            return dataKey;
        }
        catch (CryptographicException)
        {
            throw new InvalidOperationException("Wrong password.");
        }
    }

    private static byte[] DecryptPackage(byte[] encryptedPackage, byte[] key, byte[] keyDataSalt)
    {
        // Format: [8 bytes: total original size][encrypted 4096-byte chunks]
        var totalSize = BitConverter.ToInt64(encryptedPackage, 0);
        var remaining = (int)totalSize;
        var result = new byte[totalSize];
        var offset = 8; // skip size prefix
        var resultOffset = 0;
        var chunkIndex = 0;

        while (remaining > 0 && offset < encryptedPackage.Length)
        {
            var iv = DeriveIv(keyDataSalt, chunkIndex);
            var chunkLen = Math.Min(ChunkSize, encryptedPackage.Length - offset);
            var encryptedChunk = new byte[chunkLen];
            Buffer.BlockCopy(encryptedPackage, offset, encryptedChunk, 0, chunkLen);

            var decrypted = AesCbcDecrypt(encryptedChunk, key, iv);
            var toCopy = Math.Min(decrypted.Length, remaining);
            Buffer.BlockCopy(decrypted, 0, result, resultOffset, toCopy);

            resultOffset += toCopy;
            remaining -= toCopy;
            offset += chunkLen;
            chunkIndex++;
        }

        return result;
    }

    private static byte[] EncryptPackage(byte[] innerZip, byte[] key, byte[] keyDataSalt)
    {
        // Format: [8 bytes: total size][encrypted 4096-byte chunks]
        using var ms = new MemoryStream();
        ms.Write(BitConverter.GetBytes((long)innerZip.Length), 0, 8);

        var offset = 0;
        var chunkIndex = 0;
        while (offset < innerZip.Length)
        {
            var iv = DeriveIv(keyDataSalt, chunkIndex);
            var chunkLen = Math.Min(ChunkSize, innerZip.Length - offset);
            var chunk = new byte[chunkLen];
            Buffer.BlockCopy(innerZip, offset, chunk, 0, chunkLen);

            // Pad last chunk to block size
            var padded = PadToBlockSize(chunk, BlockSize);
            var encrypted = AesCbcEncrypt(padded, key, iv);
            ms.Write(encrypted, 0, encrypted.Length);

            offset += chunkLen;
            chunkIndex++;
        }

        return ms.ToArray();
    }

    private static byte[] BuildEncryptionInfoXml(
        byte[] keyDataSalt, byte[] passwordSalt, int spinCount,
        byte[] encryptedKeyValue, byte[] encryptedVerifier, byte[] encryptedVerifierHash,
        byte[] encryptedHmacKey, byte[] encryptedHmacValue)
    {
        XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
        XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        var b64 = (byte[] b) => Convert.ToBase64String(b);

        var doc = new XElement(ns + "encryption",
            new XAttribute(XNamespace.Xmlns + "p", p),

            // keyData — self-closing, all attributes
            new XElement(ns + "keyData",
                new XAttribute("saltSize", SaltSize),
                new XAttribute("blockSize", BlockSize),
                new XAttribute("keyBits", KeyBits),
                new XAttribute("hashSize", HashSize),
                new XAttribute("cipherAlgorithm", "AES"),
                new XAttribute("cipherChaining", "ChainingModeCBC"),
                new XAttribute("hashAlgorithm", "SHA512"),
                new XAttribute("saltValue", b64(keyDataSalt))),

            // dataIntegrity
            new XElement(ns + "dataIntegrity",
                new XAttribute("encryptedHmacKey", b64(encryptedHmacKey)),
                new XAttribute("encryptedHmacValue", b64(encryptedHmacValue))),

            // keyEncryptors — sibling of keyData, not child
            new XElement(ns + "keyEncryptors",
                new XElement(ns + "keyEncryptor",
                    new XAttribute("uri", p.NamespaceName),
                    // Single p:encryptedKey with ALL crypto attributes
                    new XElement(p + "encryptedKey",
                        new XAttribute("spinCount", spinCount),
                        new XAttribute("saltSize", SaltSize),
                        new XAttribute("blockSize", BlockSize),
                        new XAttribute("keyBits", KeyBits),
                        new XAttribute("hashSize", HashSize),
                        new XAttribute("cipherAlgorithm", "AES"),
                        new XAttribute("cipherChaining", "ChainingModeCBC"),
                        new XAttribute("hashAlgorithm", "SHA512"),
                        new XAttribute("saltValue", b64(passwordSalt)),
                        new XAttribute("encryptedVerifierHashInput", b64(encryptedVerifier)),
                        new XAttribute("encryptedVerifierHashValue", b64(encryptedVerifierHash)),
                        new XAttribute("encryptedKeyValue", b64(encryptedKeyValue)))))
        );

        // Prepend version + flags (8 bytes per ECMA-376)
        var header = new byte[] { 4, 0, 4, 0, 0x40, 0, 0, 0 };
        var xmlBytes = Encoding.UTF8.GetBytes(doc.ToString(SaveOptions.DisableFormatting));
        var result = new byte[header.Length + xmlBytes.Length];
        Buffer.BlockCopy(header, 0, result, 0, header.Length);
        Buffer.BlockCopy(xmlBytes, 0, result, header.Length, xmlBytes.Length);
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
