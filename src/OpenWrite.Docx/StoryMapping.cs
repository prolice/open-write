using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Docx;

/// <summary>
/// Partie de texte du package (document.xml, en-têtes, pieds, notes, commentaires) lue dans le modèle en flux.
/// La racine et ses attributs (déclarations d'espaces de noms, mc:Ignorable) sont conservés ; ses enfants sont
/// soit des conteneurs de texte (w:body, w:footnote, w:comment…), soit des îlots opaques (w:background…).
/// </summary>
public sealed class StoryPart
{
    public required string PartName { get; init; }
    /// <summary>document, header, footer, footnotes, endnotes, comments.</summary>
    public required string Kind { get; init; }
    public XDeclaration? Declaration { get; init; }
    public required XName RootName { get; init; }
    public List<XAttribute> RootAttributes { get; } = new();
    /// <summary>En-tête et pied : la racine est elle-même une suite de blocs.</summary>
    public Story? RootStory { get; set; }
    /// <summary>Enfants de la racine : <see cref="StoryHolder"/> ou <see cref="XElement"/> opaque.</summary>
    public List<object> Items { get; } = new();
    /// <summary>Partie modifiée : elle sera régénérée à l'enregistrement (les autres sont recopiées telles quelles).</summary>
    public bool Dirty { get; set; }
    /// <summary>Nœuds XML non éléments ignorés à la lecture (commentaires XML, instructions de traitement).</summary>
    public int DroppedNodes { get; set; }

    public XNamespace W => RootName.Namespace;

    public IEnumerable<Story> Stories() => RootStory is not null ? new[] { RootStory } : Items.OfType<StoryHolder>().Select(h => h.Story);
}

/// <summary>Élément qui contient une suite de blocs : w:body, w:footnote, w:endnote, w:comment.</summary>
public sealed class StoryHolder
{
    public required XName Name { get; init; }
    public List<XAttribute> Attributes { get; } = new();
    public Story Story { get; } = new();
}

/// <summary>Lecture et écriture WordprocessingML ↔ modèle en flux (prototype du lot 1).</summary>
public static class StoryMapping
{
    static readonly HashSet<string> HolderNames = new() { "body", "footnote", "endnote", "comment" };
    static readonly HashSet<string> ContainerNames = new() { "hyperlink", "ins", "del", "moveFrom", "moveTo", "smartTag", "customXml", "fldSimple", "dir", "bdo" };
    static readonly HashSet<string> ContainerHeads = new() { "smartTagPr", "customXmlPr" };

    // ================================================================ lecture

    public static StoryPart Read(XDocument xml, string partName, string kind)
    {
        var root = xml.Root ?? throw new InvalidDataException($"{partName} : document XML vide.");
        var part = new StoryPart { PartName = partName, Kind = kind, Declaration = xml.Declaration, RootName = root.Name };
        part.RootAttributes.AddRange(root.Attributes().Select(a => new XAttribute(a)));
        var w = root.Name.Namespace;
        var ctx = new Ctx(w, part);
        if (kind is "header" or "footer")
        {
            part.RootStory = new Story();
            ReadBlocks(root, part.RootStory, ctx, isBody: false);
            return part;
        }
        foreach (var node in root.Nodes())
        {
            if (node is not XElement e) { ctx.Drop(node); continue; }
            if (e.Name.Namespace == w && HolderNames.Contains(e.Name.LocalName))
            {
                var h = new StoryHolder { Name = e.Name };
                h.Attributes.AddRange(Attrs(e));
                ReadBlocks(e, h.Story, ctx, isBody: e.Name.LocalName == "body");
                part.Items.Add(h);
            }
            else part.Items.Add(new XElement(e));
        }
        return part;
    }

    sealed class Ctx(XNamespace w, StoryPart part)
    {
        public XNamespace W { get; } = w;
        public void Drop(XNode n)
        {
            // les blancs entre éléments ne portent rien ; les commentaires XML et instructions sont comptés
            if (n is XText t && string.IsNullOrWhiteSpace(t.Value)) return;
            part.DroppedNodes++;
        }
    }

    static IEnumerable<XAttribute> Attrs(XElement e) => e.Attributes().Select(a => new XAttribute(a));

    static PropertyBag Bag(string local, XElement? e, Ctx c) => new(local, c.W, e is null ? null : new XElement(e));

    static void ReadBlocks(XElement container, Story story, Ctx c, bool isBody)
    {
        var elements = container.Nodes().ToList();
        var lastElement = container.Elements().LastOrDefault();
        foreach (var node in elements)
        {
            if (node is not XElement e) { c.Drop(node); continue; }
            if (e.Name.Namespace != c.W) { story.Blocks.Add(new OpaqueBlock(new XElement(e))); continue; }
            switch (e.Name.LocalName)
            {
                case "p": story.Blocks.Add(ReadParagraph(e, c)); break;
                case "tbl": story.Blocks.Add(ReadTable(e, c)); break;
                case "sdt": story.Blocks.Add(ReadBlockSdt(e, c)); break;
                case "sectPr" when isBody && e == lastElement: story.FinalSection = Bag("sectPr", e, c); break;
                default:
                    story.Blocks.Add(Marker(e, c) is { } m ? new InlineMarkerBlock(m) : new OpaqueBlock(new XElement(e)));
                    break;
            }
        }
    }

    static Paragraph ReadParagraph(XElement e, Ctx c)
    {
        var p = new Paragraph();
        p.Attributes.AddRange(Attrs(e));
        var pPr = e.Element(c.W + "pPr");
        p.Props = Bag("pPr", pPr, c);
        ReadInlines(e.Nodes().Where(n => n != pPr), p.Inlines, c);
        return p;
    }

    static void ReadInlines(IEnumerable<XNode> nodes, List<Inline> target, Ctx c)
    {
        foreach (var node in nodes)
        {
            if (node is not XElement e) { c.Drop(node); continue; }
            if (e.Name.Namespace != c.W) { target.Add(new OpaqueInline(new XElement(e))); continue; }
            var name = e.Name.LocalName;
            if (name == "r") target.Add(ReadRun(e, c));
            else if (ContainerNames.Contains(name))
            {
                var box = new InlineContainer(e.Name);
                box.Attributes.AddRange(Attrs(e));
                var heads = e.Elements().Where(h => h.Name.Namespace == c.W && ContainerHeads.Contains(h.Name.LocalName)).ToList();
                box.Head.AddRange(heads.Select(h => new XElement(h)));
                ReadInlines(e.Nodes().Where(n => n is not XElement x || !heads.Contains(x)), box.Children, c);
                target.Add(box);
            }
            else if (name == "sdt")
            {
                var sdt = new InlineContentControl { SdtPr = Copy(e.Element(c.W + "sdtPr")), SdtEndPr = Copy(e.Element(c.W + "sdtEndPr")) };
                sdt.Attributes.AddRange(Attrs(e));
                if (e.Element(c.W + "sdtContent") is { } content)
                {
                    sdt.ContentAttributes.AddRange(Attrs(content));
                    ReadInlines(content.Nodes(), sdt.Content, c);
                }
                target.Add(sdt);
            }
            else target.Add(Marker(e, c) ?? (Inline)new OpaqueInline(new XElement(e)));
        }
    }

    static Run ReadRun(XElement e, Ctx c)
    {
        var r = new Run();
        r.Attributes.AddRange(Attrs(e));
        var rPr = e.Element(c.W + "rPr");
        r.Props = Bag("rPr", rPr, c);
        foreach (var node in e.Nodes())
        {
            if (node == rPr) continue;
            if (node is not XElement x) { c.Drop(node); continue; }
            RunContent item = x.Name.Namespace != c.W ? new OpaqueRunContent(new XElement(x)) : x.Name.LocalName switch
            {
                "t" => new TextContent { Value = x.Value, LoadedValue = x.Value },
                "delText" => new TextContent { Value = x.Value, LoadedValue = x.Value, Deleted = true },
                "fldChar" => FieldCharOf(x, c),
                "instrText" => new FieldCode { Value = x.Value, LoadedValue = x.Value },
                "delInstrText" => new FieldCode { Value = x.Value, LoadedValue = x.Value, Deleted = true },
                "tab" => new TabContent(),
                "br" => new BreakContent(),
                "cr" => new BreakContent { CarriageReturn = true },
                _ => new OpaqueRunContent(new XElement(x)),
            };
            if (item is not OpaqueRunContent and not FieldChar) item.Attributes.AddRange(Attrs(x));
            r.Content.Add(item);
        }
        return r;
    }

    static Table ReadTable(XElement e, Ctx c)
    {
        var t = new Table();
        t.Attributes.AddRange(Attrs(e));
        foreach (var node in e.Nodes())
        {
            if (node is not XElement x) { c.Drop(node); continue; }
            if (x.Name == c.W + "tblPr") t.Props = Bag("tblPr", x, c);
            else if (x.Name == c.W + "tblGrid") t.Grid = new XElement(x);
            else if (x.Name == c.W + "tr") t.Items.Add(ReadRow(x, c));
            else t.Items.Add(Marker(x, c) ?? (object)new XElement(x));
        }
        return t;
    }

    static TableRow ReadRow(XElement e, Ctx c)
    {
        var row = new TableRow();
        row.Attributes.AddRange(Attrs(e));
        foreach (var node in e.Nodes())
        {
            if (node is not XElement x) { c.Drop(node); continue; }
            if (x.Name == c.W + "tblPrEx") row.PropsException = Bag("tblPrEx", x, c);
            else if (x.Name == c.W + "trPr") row.Props = Bag("trPr", x, c);
            else if (x.Name == c.W + "tc")
            {
                var cell = new TableCell();
                cell.Attributes.AddRange(Attrs(x));
                var tcPr = x.Element(c.W + "tcPr");
                cell.Props = Bag("tcPr", tcPr, c);
                var holder = new XElement(x);
                holder.Element(c.W + "tcPr")?.Remove();
                ReadBlocks(holder, cell.Content, c, isBody: false);
                row.Items.Add(cell);
            }
            else row.Items.Add(Marker(x, c) ?? (object)new XElement(x));
        }
        return row;
    }

    static BlockContentControl ReadBlockSdt(XElement e, Ctx c)
    {
        var sdt = new BlockContentControl { SdtPr = Copy(e.Element(c.W + "sdtPr")), SdtEndPr = Copy(e.Element(c.W + "sdtEndPr")) };
        sdt.Attributes.AddRange(Attrs(e));
        if (e.Element(c.W + "sdtContent") is { } content)
        {
            sdt.ContentAttributes.AddRange(Attrs(content));
            ReadBlocks(content, sdt.Content, c, isBody: false);
        }
        return sdt;
    }

    static XElement? Copy(XElement? e) => e is null ? null : new XElement(e);

    /// <summary>Marqueur de signet typé ; null pour tout autre élément.</summary>
    static Inline? Marker(XElement e, Ctx c)
    {
        if (e.Name.Namespace != c.W) return null;
        Inline? m = e.Name.LocalName switch
        {
            "bookmarkStart" => new BookmarkStart(c.W),
            "bookmarkEnd" => new BookmarkEnd(c.W),
            _ => null,
        };
        if (m is not null)
        {
            if (e.HasElements || e.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value))) return null;   // forme inattendue : opaque
            m.Attributes.AddRange(Attrs(e));
        }
        return m;
    }

    static FieldChar FieldCharOf(XElement x, Ctx c)
    {
        var fc = new FieldChar(c.W);
        fc.Attributes.AddRange(Attrs(x));
        fc.Children.AddRange(x.Elements().Select(ch => new XElement(ch)));
        return fc;
    }

    // ================================================================ écriture

    public static XDocument Write(StoryPart part)
    {
        var w = part.W;
        var root = new XElement(part.RootName, part.RootAttributes.Select(a => new XAttribute(a)));
        if (part.RootStory is not null) root.Add(WriteBlocks(part.RootStory, w));
        else
            foreach (var item in part.Items)
                root.Add(item switch
                {
                    StoryHolder h => new XElement(h.Name, h.Attributes.Select(a => new XAttribute(a)), WriteBlocks(h.Story, w)),
                    XElement x => new XElement(x),
                    _ => throw new InvalidOperationException(),
                });
        return new XDocument(part.Declaration ?? new XDeclaration("1.0", "UTF-8", "yes"), root);
    }

    static IEnumerable<XAttribute> Copy(List<XAttribute> attrs) => attrs.Select(a => new XAttribute(a));

    static XElement? Out(PropertyBag? bag) => bag?.Element is { } e ? new XElement(e) : null;

    static IEnumerable<XElement> WriteBlocks(Story story, XNamespace w)
    {
        foreach (var b in story.Blocks)
            yield return b switch
            {
                Paragraph p => new XElement(w + "p", Copy(p.Attributes), Out(p.Props), p.Inlines.Select(i => WriteInline(i, w))),
                Table t => WriteTable(t, w),
                BlockContentControl s => new XElement(w + "sdt", Copy(s.Attributes), Copy(s.SdtPr), Copy(s.SdtEndPr),
                    new XElement(w + "sdtContent", Copy(s.ContentAttributes), WriteBlocks(s.Content, w))),
                OpaqueBlock o => new XElement(o.Xml),
                InlineMarkerBlock mb => WriteInline(mb.Marker, w),
                _ => throw new InvalidOperationException(b.GetType().Name),
            };
        if (Out(story.FinalSection) is { } sect) yield return sect;
    }

    static XElement WriteTable(Table t, XNamespace w) =>
        new(w + "tbl", Copy(t.Attributes), Out(t.Props), Copy(t.Grid),
            t.Items.Select(item => item switch
            {
                TableRow r => new XElement(w + "tr", Copy(r.Attributes), Out(r.PropsException), Out(r.Props),
                    r.Items.Select(ci => ci switch
                    {
                        TableCell cell => new XElement(w + "tc", Copy(cell.Attributes), Out(cell.Props), WriteBlocks(cell.Content, w)),
                        XElement x => new XElement(x),
                        Inline m => WriteInline(m, w),
                        _ => throw new InvalidOperationException(),
                    })),
                XElement x => new XElement(x),
                Inline m => WriteInline(m, w),
                _ => throw new InvalidOperationException(),
            }));

    static XElement WriteInline(Inline i, XNamespace w) => i switch
    {
        Run r => new XElement(w + "r", Copy(r.Attributes), Out(r.Props), r.Content.Select(c => WriteRunContent(c, w))),
        InlineContainer box => new XElement(box.Name, Copy(box.Attributes), box.Head.Select(h => new XElement(h)), box.Children.Select(ch => WriteInline(ch, w))),
        InlineContentControl s => new XElement(w + "sdt", Copy(s.Attributes), Copy(s.SdtPr), Copy(s.SdtEndPr),
            new XElement(w + "sdtContent", Copy(s.ContentAttributes), s.Content.Select(ch => WriteInline(ch, w)))),
        OpaqueInline o => new XElement(o.Xml),
        BookmarkStart b => new XElement(w + "bookmarkStart", Copy(b.Attributes)),
        BookmarkEnd b => new XElement(w + "bookmarkEnd", Copy(b.Attributes)),
        _ => throw new InvalidOperationException(i.GetType().Name),
    };

    static XElement WriteRunContent(RunContent c, XNamespace w)
    {
        if (c is OpaqueRunContent o) return new XElement(o.Xml);
        if (c is FieldChar fc) return new XElement(w + "fldChar", Copy(fc.Attributes), fc.Children.Select(ch => new XElement(ch)));
        var e = c switch
        {
            TextContent t => new XElement(w + (t.Deleted ? "delText" : "t"), t.Value),
            FieldCode f => new XElement(w + (f.Deleted ? "delInstrText" : "instrText"), f.Value),
            TabContent => new XElement(w + "tab"),
            BreakContent b => new XElement(w + (b.CarriageReturn ? "cr" : "br")),
            _ => throw new InvalidOperationException(c.GetType().Name),
        };
        e.Add(Copy(c.Attributes));
        // texte avec blancs XML en tête ou en fin : xml:space="preserve" obligatoire, sinon Word les supprime.
        // Blancs au sens XML seulement (espace, tabulation, retours) : l'espace insécable n'en est pas un (calibration
        // corpus : « \u00A0:… » enregistré par Word sans xml:space).
        var (value, loaded) = c switch { TextContent t2 => (t2.Value, t2.LoadedValue), FieldCode f2 => (f2.Value, f2.LoadedValue), _ => ("", "") };
        if (value != loaded && value.Length > 0 && (IsXmlSpace(value[0]) || IsXmlSpace(value[^1]))
            && e.Attribute(XNamespace.Xml + "space") is null)
            e.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        return e;
    }

    static bool IsXmlSpace(char ch) => ch is ' ' or '\t' or '\n' or '\r';
}
