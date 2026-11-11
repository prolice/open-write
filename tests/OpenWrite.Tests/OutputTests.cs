// Repris d'OpenPub 2.1.4 (tests/OpenPub.Tests/OutputTests.cs) : choix des pages, livret, recto verso manuel.
using OpenWrite.Core.Output;

namespace OpenWrite.Tests;

public class OutputTests
{
    [Fact]
    public void Livret_de_8_pages_dans_l_ordre_de_pliage()
    {
        var sides = Booklet.Sides(8);
        Assert.Equal(new[] { (8, 1), (2, 7), (6, 3), (4, 5) }, sides.Select(s => (s.Left, s.Right)));
        Assert.Equal(new[] { true, false, true, false }, sides.Select(s => s.Front));
        Assert.Equal(new[] { 1, 1, 2, 2 }, sides.Select(s => s.Sheet));
    }

    [Fact]
    public void Livret_complete_par_des_pages_blanches()
    {
        Assert.Equal(4, Booklet.PaddedCount(1));
        Assert.Equal(8, Booklet.PaddedCount(5));
        Assert.Equal(12, Booklet.PaddedCount(12));
        // 6 pages -> 8 : les pages 7 et 8 sont blanches (0)
        var sides = Booklet.Sides(6);
        Assert.Equal(new[] { (0, 1), (2, 0), (6, 3), (4, 5) }, sides.Select(s => (s.Left, s.Right)));
        // chaque page reelle apparait une fois exactement
        var pages = Booklet.Sides(10).SelectMany(s => new[] { s.Left, s.Right }).Where(p => p > 0).OrderBy(p => p);
        Assert.Equal(Enumerable.Range(1, 10), pages);
    }

    [Fact]
    public void Recto_verso_manuel_separe_rectos_et_versos()
    {
        var (fronts, backs) = ManualDuplex.Split(new[] { 1, 2, 3, 4, 5 }, 0);
        Assert.Equal(new[] { 1, 3, 5 }, fronts);
        Assert.Equal(new[] { 2, 4, 0 }, backs);   // verso blanc pour la derniere feuille
        var (_, reversed) = ManualDuplex.Split(new[] { 1, 2, 3, 4, 5 }, 0, reverseBacks: true);
        Assert.Equal(new[] { 0, 4, 2 }, reversed);
    }

    [Fact]
    public void Choix_des_pages_plages_et_erreurs()
    {
        Assert.Equal(new[] { 1, 2, 3, 4 }, PageRange.Parse("", 4, out _));
        Assert.Equal(new[] { 1, 2, 3, 4 }, PageRange.Parse("Toutes", 4, out _));
        Assert.Equal(new[] { 1, 2, 3, 5 }, PageRange.Parse("1-3; 5", 6, out _));
        Assert.Equal(new[] { 4, 5, 6, 2 }, PageRange.Parse("4-, 2, 5", 6, out _));
        Assert.Equal(new[] { 2, 3 }, PageRange.Parse("3-2", 6, out _));
        Assert.Null(PageRange.Parse("1-9", 6, out var err));
        Assert.Contains("6 page", err);
        Assert.Null(PageRange.Parse("deux", 6, out err));
        Assert.Contains("deux", err);
    }
}
