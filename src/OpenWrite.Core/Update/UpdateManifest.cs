// Repris d'OpenPub 2.1.3 (src/OpenPub.Core/Update/UpdateManifest.cs), adapté à OpenWrite.
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OpenWrite.Core.Update;

/// <summary>
/// Description de la derniere version publiee (fichier latest.json sur le site de mises a jour).
/// Le contenu utile (version, lien, empreinte, taille) est signe avec une cle privee qui ne quitte pas
/// le poste de publication ; l'application verifie la signature avec la cle publique qu'elle embarque.
/// </summary>
public sealed class UpdateManifest
{
    public string Version { get; set; } = "0.0.0";
    /// <summary>Adresse de l'installeur MSI (absolue, ou relative a latest.json).</summary>
    public string Url { get; set; } = "";
    /// <summary>Empreinte SHA-256 de l'installeur, en hexadecimal minuscule.</summary>
    public string Sha256 { get; set; } = "";
    public long Size { get; set; }
    public string? Notes { get; set; }
    public DateTime Published { get; set; }
    /// <summary>Signature RSA-SHA256 (PKCS#1) de <see cref="Payload"/>, en base 64.</summary>
    public string Signature { get; set; } = "";

    static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>Texte signe : tout ce qui determine le fichier installe.</summary>
    [JsonIgnore]
    public string Payload => $"openwrite-update/1\n{Version}\n{Url}\n{Sha256.ToLowerInvariant()}\n{Size}";

    public string ToJson() => JsonSerializer.Serialize(this, Json);
    public static UpdateManifest Parse(string json) => JsonSerializer.Deserialize<UpdateManifest>(json, Json) ?? throw new InvalidDataException("latest.json vide");

    [JsonIgnore] public System.Version ParsedVersion => System.Version.TryParse(Version, out var v) ? v : new System.Version(0, 0);
}

/// <summary>Cles et signatures des mises a jour.</summary>
public static class UpdateSigner
{
    /// <summary>Cree une paire de cles ; renvoie (cle privee PEM, cle publique en base 64).</summary>
    public static (string PrivatePem, string PublicKey) GenerateKey()
    {
        using var rsa = RSA.Create(3072);
        return (rsa.ExportRSAPrivateKeyPem(), Convert.ToBase64String(rsa.ExportSubjectPublicKeyInfo()));
    }

    public static string PublicKeyOf(string privatePem)
    {
        using var rsa = RSA.Create();
        rsa.ImportFromPem(privatePem);
        return Convert.ToBase64String(rsa.ExportSubjectPublicKeyInfo());
    }

    public static void Sign(UpdateManifest m, string privatePem)
    {
        using var rsa = RSA.Create();
        rsa.ImportFromPem(privatePem);
        var sig = rsa.SignData(Encoding.UTF8.GetBytes(m.Payload), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        m.Signature = Convert.ToBase64String(sig);
    }

    public static bool Verify(UpdateManifest m, string publicKeyBase64)
    {
        if (string.IsNullOrWhiteSpace(m.Signature) || string.IsNullOrWhiteSpace(publicKeyBase64)) return false;
        try
        {
            using var rsa = RSA.Create();
            rsa.ImportSubjectPublicKeyInfo(Convert.FromBase64String(publicKeyBase64), out _);
            return rsa.VerifyData(Encoding.UTF8.GetBytes(m.Payload), Convert.FromBase64String(m.Signature),
                HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        }
        catch (FormatException) { return false; }
        catch (CryptographicException) { return false; }
    }

    public static string Sha256Of(string file)
    {
        using var fs = File.OpenRead(file);
        return Convert.ToHexString(SHA256.HashData(fs)).ToLowerInvariant();
    }
}
