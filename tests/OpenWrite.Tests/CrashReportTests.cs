// Repris d'OpenPub 2.1.3 (tests/OpenPub.Tests/CrashReportTests.cs).
using OpenWrite.Core.Diagnostics;

namespace OpenWrite.Tests;

/// <summary>Rapports d'erreur : anonymisation (depot public), signature stable, file d'attente locale.</summary>
public class CrashReportTests : IDisposable
{
    readonly string _dir = Path.Combine(Path.GetTempPath(), "openwrite-tests-" + Guid.NewGuid().ToString("N")[..8]);

    public void Dispose()
    {
        try { Directory.Delete(_dir, true); } catch { }
    }

    [Fact]
    public void Anonymise_chemins_noms_de_fichier_et_e_mails()
    {
        var s = CrashReport.Scrub(@"Impossible d'ouvrir C:\Users\jdupont\OneDrive - Ecole\Bulletin de Lucie.docx (contact jean.dupont@ecole.be)");
        Assert.DoesNotContain("jdupont", s);
        Assert.DoesNotContain("Lucie", s);
        Assert.DoesNotContain("@ecole", s);
        Assert.Contains("%USERPROFILE%", s);
        Assert.Contains("….docx", s);
    }

    [Fact]
    public void Anonymise_le_texte_cite_mais_garde_les_identifiants()
    {
        var s = CrashReport.Scrub("Valeur 'Bonjour à tous les parents' invalide pour 'FontSize'");
        Assert.DoesNotContain("parents", s);
        Assert.Contains("'FontSize'", s);
    }

    [Fact]
    public void Chemins_des_sources_reduits_au_nom_du_fichier()
    {
        var s = CrashReport.Scrub(@"   at OpenWrite.App.X.Y() in C:\Users\jdupont\src\OpenWrite.App\X.cs:line 12");
        Assert.Equal("   at OpenWrite.App.X.Y() in X.cs:line 12", s);
    }

    static Exception Throw(Action a)
    {
        try { a(); } catch (Exception e) { return e; }
        throw new InvalidOperationException("pas d'exception");
    }

    static void Faute(string msg) => throw new InvalidOperationException(msg);

    [Fact]
    public void Signature_stable_independante_du_message_et_des_lignes()
    {
        // meme lieu d'erreur (meme lambda), messages et versions differents
        static Exception Meme(string m) => Throw(() => Faute(m));
        var a = CrashReport.From(Meme("un"), "interface", false, "2.1.3");
        var b = CrashReport.From(Meme("deux"), "interface", false, "2.1.4");
        Assert.Equal(a.Signature, b.Signature);
        Assert.Matches("^[0-9a-f]{12}$", a.Signature);
        Assert.Equal("System.InvalidOperationException", a.Type);
        var c = CrashReport.From(Throw(() => throw new ArgumentException("x")), "interface", false, "2.1.3");
        Assert.NotEqual(a.Signature, c.Signature);
        // insensible aux numeros de ligne
        Assert.Equal(CrashReport.SignatureOf("T", "   at OpenWrite.A.B() in A.cs:line 1"), CrashReport.SignatureOf("T", "   at OpenWrite.A.B() in A.cs:line 99"));
    }

    [Fact]
    public void Exception_enveloppee_ramenee_a_la_cause()
    {
        var r = CrashReport.From(new AggregateException(Throw(() => Faute("x"))), "tâche", false, "2.1.3");
        Assert.Equal("System.InvalidOperationException", r.Type);
    }

    [Fact]
    public void File_d_attente_sur_disque_et_limitee()
    {
        var store = new CrashStore(_dir);
        for (int i = 0; i < CrashStore.MaxPending + 5; i++)
        {
            var r = CrashReport.From(Throw(() => Faute("x")), "interface", false, "2.1.3");
            r.TimeUtc = new DateTime(2026, 11, 4, 10, 0, 0, DateTimeKind.Utc).AddSeconds(i);
            store.Save(r);
        }
        File.WriteAllText(Path.Combine(_dir, "0-illisible.json"), "{pas du json");
        var pending = store.Pending();
        Assert.Equal(CrashStore.MaxPending, pending.Count);
        Assert.Equal(5, pending[0].Report.TimeUtc.Second);   // les plus anciens ont ete supprimes
        Assert.False(File.Exists(Path.Combine(_dir, "0-illisible.json")));
        Assert.Equal(pending[0].Report.Signature, CrashReport.Parse(File.ReadAllText(pending[0].Path))!.Signature);
    }
}
