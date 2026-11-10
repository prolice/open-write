using OpenWrite.Core.Model;

namespace OpenWrite.Layout;

/// <summary>
/// Position du point d'insertion dans le texte affiché d'un paragraphe (lot 3) : avant le caractère <see cref="Index"/>
/// de <see cref="ParagraphContent.Text"/> (0 à longueur). <see cref="AtLineEnd"/> : à la fin d'une ligne coupée plutôt
/// qu'au début de la suivante (même index ; touche Fin, clic après la fin d'une ligne).
/// </summary>
public readonly record struct TextPoint(Paragraph Paragraph, int Index, bool AtLineEnd = false);

/// <summary>Ligne placée, désignée par la page et son rang dans <see cref="Page.Lines"/> (ordre du document).</summary>
public readonly record struct LineRef(int Page, int Line) : IComparable<LineRef>
{
    public int CompareTo(LineRef o) => Page != o.Page ? Page.CompareTo(o.Page) : Line.CompareTo(o.Line);
}

/// <summary>
/// Géométrie du point d'insertion : correspondance texte affiché ↔ positions du modèle, abscisse d'un caractère sur
/// sa ligne, caractère le plus proche d'une abscisse, mots. Les abscisses sont relatives à <see cref="PlacedLine.X"/>.
/// </summary>
public static class CaretGeometry
{
    // ------------------------------------------------------------------ texte affiché ↔ modèle

    /// <summary>
    /// Décalage dans le modèle du point d'insertion placé avant le caractère affiché <paramref name="i"/> : juste après le
    /// caractère qui précède (la frappe prend sa mise en forme, comme dans Word), ou 0 en tête de paragraphe.
    /// </summary>
    public static int ModelOffset(ParagraphContent c, int i) => i <= 0 || c.ModelEnd.Length == 0 ? 0 : c.ModelEnd[Math.Min(i, c.ModelEnd.Length) - 1];

    /// <summary>Plage du modèle couverte par le caractère affiché <paramref name="i"/> (à supprimer par Retour arrière / Suppr).</summary>
    public static (int Start, int End) ModelRange(ParagraphContent c, int i) => (c.ModelStart[i], c.ModelEnd[i]);

    /// <summary>Caractère affiché correspondant à un décalage du modèle : dernier point d'insertion qui ne le dépasse pas.</summary>
    public static int DisplayIndex(ParagraphContent c, int offset)
    {
        int n = c.Text.Length, best = 0;
        for (int i = 1; i <= n; i++)
            if (ModelOffset(c, i) <= offset) best = i; else break;
        return best;
    }

    // ------------------------------------------------------------------ lignes

    /// <summary>Ligne du paragraphe qui contient le point d'insertion.</summary>
    public static int LineOf(ParagraphLayout pl, int index, bool atLineEnd)
    {
        var lines = pl.Lines;
        for (int k = 0; k < lines.Count; k++)
        {
            var l = lines[k];
            if (index < l.Start) return Math.Max(0, k - 1);
            if (index < l.End) return k;
            // fin de ligne : sur cette ligne si c'est la dernière, si elle finit par un saut, ou si demandé
            if (index == l.End && (k == lines.Count - 1 || l.EndKind != LineEnd.Wrap || atLineEnd || lines[k + 1].Start > index)) return k;
        }
        return Math.Max(0, lines.Count - 1);
    }

    /// <summary>Début et fin des positions de la ligne (la fin d'une ligne coupée = début de la suivante).</summary>
    public static (int Start, int End) LineSpan(ParagraphLayout pl, int line)
    {
        var l = pl.Lines[line];
        return (l.Start, l.End);
    }

    /// <summary>Abscisse (relative à la ligne placée) du point d'insertion avant le caractère <paramref name="index"/>.</summary>
    public static double X(ParagraphLayout pl, int line, int index)
    {
        var box = pl.Lines[line];
        var c = pl.Content;
        if (box.Pieces.Count == 0) return EmptyLineX(pl, line);
        double shift = 0;
        double lastEnd = box.Pieces[0].X;
        int lastEndIndex = -1;
        foreach (var piece in box.Pieces)
        {
            int end = piece.Start + piece.Length;
            // avant une tabulation : fin du morceau précédent ; après : début du morceau qui suit le taquet
            if (index < piece.Start) return index <= lastEndIndex ? lastEnd : piece.X + shift;
            double x = piece.X + shift;
            int spaces = 0;
            for (int k = piece.Start; k < end; k++)
            {
                if (k == index) return x;
                x += Advance(c, k) + (c.Text[k] == ' ' ? box.ExtraPerSpace : 0);
                if (c.Text[k] == ' ') spaces++;
            }
            shift += spaces * box.ExtraPerSpace;
            lastEnd = x;
            lastEndIndex = end;
            if (index == end && end == box.End) return x;
        }
        return lastEnd;
    }

    /// <summary>Position du texte d'une ligne vide (paragraphe vide) : retrait et alignement du paragraphe.</summary>
    static double EmptyLineX(ParagraphLayout pl, int line)
    {
        var f = pl.Format;
        double left = f.IndentLeft / 20.0 + (line == 0 ? f.IndentFirstLine / 20.0 : 0), right = pl.Width - f.IndentRight / 20.0;
        if (line == 0 && pl.Marker is { } m) left = m.X + m.Style.Face.Width(m.Text, m.Style.Size) + 2;
        return f.Alignment switch
        {
            "center" => (left + right) / 2,
            "right" => right,
            _ => left,
        };
    }

    /// <summary>Point d'insertion le plus proche d'une abscisse sur une ligne (frontière de caractère la plus proche).</summary>
    public static int IndexAtX(ParagraphLayout pl, int line, double x)
    {
        var box = pl.Lines[line];
        int end = box.End;
        // espaces de fin d'une ligne coupée : le point d'insertion reste avant le mot suivant (fin de ligne)
        int best = box.Start;
        double bestD = double.MaxValue;
        for (int i = box.Start; i <= end; i++)
        {
            double d = Math.Abs(X(pl, line, i) - x);
            if (d < bestD) { bestD = d; best = i; }
        }
        return best;
    }

    static double Advance(ParagraphContent c, int k)
    {
        if (c.Text[k] == Atom.Object && c.Objects.TryGetValue(k, out var o)) return o.W;
        var st = c.Styles[c.StyleOf[k]];
        return st.Face.Advance(c.Text[k]) * st.Size + st.Tracking;
    }

    // ------------------------------------------------------------------ mots

    public static bool IsWordChar(char ch) => char.IsLetterOrDigit(ch) || ch is '_' or '’' or '\'' or '-' or '­';

    /// <summary>Début du mot suivant (Ctrl+Droite) : fin du mot puis des espaces qui le suivent ; n en fin de paragraphe.</summary>
    public static int NextWord(string text, int i)
    {
        int n = text.Length;
        if (i >= n) return n;
        if (IsWordChar(text[i])) while (i < n && IsWordChar(text[i])) i++;
        else if (!char.IsWhiteSpace(text[i])) i++;
        while (i < n && text[i] == ' ') i++;
        return i;
    }

    /// <summary>Début du mot précédent (Ctrl+Gauche).</summary>
    public static int PreviousWord(string text, int i)
    {
        if (i <= 0) return 0;
        i--;
        while (i > 0 && text[i] == ' ') i--;
        if (IsWordChar(text[i])) while (i > 0 && IsWordChar(text[i - 1])) i--;
        return i;
    }

    /// <summary>Mot sous un caractère (double-clic) : le mot et les espaces qui le suivent, comme Word.</summary>
    public static (int Start, int End) WordAt(string text, int i)
    {
        int n = text.Length;
        if (n == 0) return (0, 0);
        i = Math.Clamp(i, 0, n - 1);
        if (!IsWordChar(text[i])) return (i, Math.Min(n, i + 1));
        int s = i, e = i;
        while (s > 0 && IsWordChar(text[s - 1])) s--;
        while (e < n && IsWordChar(text[e])) e++;
        while (e < n && text[e] == ' ') e++;
        return (s, e);
    }
}
