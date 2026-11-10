// Repris d'OpenPub 2.1.3 (tests/OpenPub.Tests/UpdateTests.cs).
using OpenWrite.Core.Update;

namespace OpenWrite.Tests;

public class UpdateTests
{
    static UpdateManifest Sample() => new()
    {
        Version = "1.2.0", Url = "OpenWrite-1.2.0.msi", Sha256 = new string('a', 64), Size = 123456, Notes = "Test",
    };

    [Fact]
    public void Signature_valide_acceptee()
    {
        var (priv, pub) = UpdateSigner.GenerateKey();
        var m = Sample();
        UpdateSigner.Sign(m, priv);
        var reread = UpdateManifest.Parse(m.ToJson());
        Assert.True(UpdateSigner.Verify(reread, pub));
        Assert.Equal(pub, UpdateSigner.PublicKeyOf(priv));
    }

    [Theory]
    [InlineData("version")]
    [InlineData("url")]
    [InlineData("sha")]
    [InlineData("size")]
    public void Manifeste_modifie_rejete(string champ)
    {
        var (priv, pub) = UpdateSigner.GenerateKey();
        var m = Sample();
        UpdateSigner.Sign(m, priv);
        switch (champ)
        {
            case "version": m.Version = "9.9.9"; break;
            case "url": m.Url = "https://pirate.example/x.msi"; break;
            case "sha": m.Sha256 = new string('b', 64); break;
            case "size": m.Size++; break;
        }
        Assert.False(UpdateSigner.Verify(m, pub));
    }

    [Fact]
    public void Autre_cle_rejetee()
    {
        var (priv, _) = UpdateSigner.GenerateKey();
        var (_, otherPub) = UpdateSigner.GenerateKey();
        var m = Sample();
        UpdateSigner.Sign(m, priv);
        Assert.False(UpdateSigner.Verify(m, otherPub));
    }

    [Fact]
    public void Les_notes_ne_sont_pas_signees_mais_le_reste_oui()
    {
        var (priv, pub) = UpdateSigner.GenerateKey();
        var m = Sample();
        UpdateSigner.Sign(m, priv);
        m.Notes = "Autre texte";
        Assert.True(UpdateSigner.Verify(m, pub));      // les notes ne determinent pas le fichier installe
    }
}
