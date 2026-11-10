using OpenWrite.Core.Io;
using OpenWrite.Layout;

namespace OpenWrite.Tests;

public class RecoveryTests : IDisposable
{
    readonly string _dir = Path.Combine(Path.GetTempPath(), "openwrite-tests-" + Guid.NewGuid().ToString("N")[..8]);

    public void Dispose()
    {
        try { Directory.Delete(_dir, true); } catch { }
    }

    [Fact]
    public void Copie_de_secours_puis_recuperation()
    {
        var store = new RecoveryStore(Path.Combine(_dir, "rec"));
        var doc = Path.Combine(_dir, "lettre.docx");
        var id = RecoveryStore.IdFor(doc, "s1");
        Assert.Equal(id, RecoveryStore.IdFor(doc.ToUpperInvariant(), "autre"));   // même fichier : même copie
        store.Save(id, doc, "lettre.docx", p => File.WriteAllText(p, "v1"));
        store.Save(id, doc, "lettre.docx", p => File.WriteAllText(p, "v2"));    // remplace la copie précédente

        var entry = Assert.Single(store.List());
        Assert.Equal(Path.GetFullPath(doc), entry.OriginalPath);
        Assert.Equal("lettre.docx", entry.Title);
        Assert.Equal("v2", File.ReadAllText(entry.CopyPath));

        store.Delete(entry.Id);
        Assert.Empty(store.List());
        Assert.StartsWith("sans-titre-s1", RecoveryStore.IdFor(null, "s1"));
    }

    [Fact]
    public void Une_erreur_pendant_la_copie_garde_la_precedente()
    {
        var store = new RecoveryStore(Path.Combine(_dir, "rec"));
        store.Save("x", null, "Sans titre", p => File.WriteAllText(p, "bonne"));
        Assert.Throws<InvalidOperationException>(() =>
            store.Save("x", null, "Sans titre", p => { File.WriteAllText(p, "moitié"); throw new InvalidOperationException(); }));
        Assert.Equal("bonne", File.ReadAllText(Assert.Single(store.List()).CopyPath));
        Assert.Empty(Directory.GetFiles(store.Dir, "*.tmp"));   // aucun fichier temporaire laissé
    }

    [Fact]
    public void Enregistrement_sur_garde_la_version_precedente()
    {
        Directory.CreateDirectory(_dir);
        var f = Path.Combine(_dir, "doc.docx");
        SafeFile.Write(f, p => File.WriteAllText(p, "un"), keepBackup: true);
        Assert.False(File.Exists(f + ".bak"));             // premier enregistrement : rien à garder
        SafeFile.Write(f, p => File.WriteAllText(p, "deux"), keepBackup: true);
        Assert.Equal("deux", File.ReadAllText(f));
        Assert.Equal("un", File.ReadAllText(f + ".bak"));
        SafeFile.Write(f, p => File.WriteAllText(p, "trois"), keepBackup: false);
        Assert.Equal("un", File.ReadAllText(f + ".bak"));  // sans option : .bak inchangé
    }
}

public class HyphenatorTests
{
    [Theory]
    [InlineData("compétitions", "com-pé-ti-tions")]
    [InlineData("malheureusement", "mal-heu-reu-se-ment")]
    [InlineData("rentrée", "ren-trée")]
    [InlineData("chat", "chat")]                // mot court : pas de césure
    public void Cesure_francaise(string word, string expected) =>
        Assert.Equal(expected, Hyphenator.French.Insert(word));
}
