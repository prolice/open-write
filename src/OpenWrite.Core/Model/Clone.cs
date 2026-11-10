using System.Xml.Linq;

namespace OpenWrite.Core.Model;

/// <summary>
/// Copie profonde des nœuds du modèle (journal d'annulation : copie « avant » et « après » des blocs touchés).
/// Tout est copié, y compris les valeurs lues (<see cref="TextContent.LoadedValue"/>) : annuler redonne exactement
/// le XML d'origine.
/// </summary>
public static class FlowClone
{
    static void CopyAttrs(List<XAttribute> from, List<XAttribute> to) => to.AddRange(from.Select(x => new XAttribute(x)));
    static XElement? X(XElement? e) => e is null ? null : new XElement(e);

    public static Story Clone(this Story s)
    {
        var c = new Story { FinalSection = s.FinalSection?.Clone() };
        c.Blocks.AddRange(s.Blocks.Select(Clone));
        return c;
    }

    static void CopyContent(Story from, Story to)
    {
        to.FinalSection = from.FinalSection?.Clone();
        to.Blocks.AddRange(from.Blocks.Select(Clone));
    }

    public static Block Clone(this Block b)
    {
        Block c = b switch
        {
            Paragraph p => CloneParagraph(p),
            Table t => CloneTable(t),
            BlockContentControl s => CloneSdt(s),
            OpaqueBlock o => new OpaqueBlock(new XElement(o.Xml)),
            InlineMarkerBlock m => new InlineMarkerBlock(m.Marker.Clone()),
            _ => throw new InvalidOperationException(b.GetType().Name),
        };
        if (b is not Paragraph and not Table and not BlockContentControl) CopyAttrs(b.Attributes, c.Attributes);
        return c;
    }

    static Paragraph CloneParagraph(Paragraph p)
    {
        var c = new Paragraph { Props = p.Props.Clone() };
        CopyAttrs(p.Attributes, c.Attributes);
        c.Inlines.AddRange(p.Inlines.Select(Clone));
        return c;
    }

    static Table CloneTable(Table t)
    {
        var c = new Table { Props = t.Props.Clone(), Grid = X(t.Grid) };
        CopyAttrs(t.Attributes, c.Attributes);
        foreach (var item in t.Items)
            c.Items.Add(item switch
            {
                TableRow r => CloneRow(r),
                Inline m => m.Clone(),
                XElement x => new XElement(x),
                _ => throw new InvalidOperationException(item.GetType().Name),
            });
        return c;
    }

    static TableRow CloneRow(TableRow r)
    {
        var c = new TableRow { Props = r.Props.Clone(), PropsException = r.PropsException?.Clone() };
        CopyAttrs(r.Attributes, c.Attributes);
        foreach (var item in r.Items)
        {
            if (item is TableCell cell)
            {
                var cc = new TableCell { Props = cell.Props.Clone() };
                CopyAttrs(cell.Attributes, cc.Attributes);
                CopyContent(cell.Content, cc.Content);
                c.Items.Add(cc);
            }
            else c.Items.Add(item switch
            {
                Inline m => m.Clone(),
                XElement x => new XElement(x),
                _ => throw new InvalidOperationException(item.GetType().Name),
            });
        }
        return c;
    }

    static BlockContentControl CloneSdt(BlockContentControl s)
    {
        var c = new BlockContentControl { SdtPr = X(s.SdtPr), SdtEndPr = X(s.SdtEndPr) };
        CopyAttrs(s.Attributes, c.Attributes);
        CopyAttrs(s.ContentAttributes, c.ContentAttributes);
        CopyContent(s.Content, c.Content);
        return c;
    }

    public static Inline Clone(this Inline i)
    {
        Inline c;
        switch (i)
        {
            case Run r:
                var rc = new Run { Props = r.Props.Clone() };
                rc.Content.AddRange(r.Content.Select(Clone));
                c = rc;
                break;
            case InlineContainer box:
                var bc = new InlineContainer(box.Name);
                bc.Head.AddRange(box.Head.Select(h => new XElement(h)));
                bc.Children.AddRange(box.Children.Select(Clone));
                c = bc;
                break;
            case InlineContentControl s:
                var sc = new InlineContentControl { SdtPr = X(s.SdtPr), SdtEndPr = X(s.SdtEndPr) };
                CopyAttrs(s.ContentAttributes, sc.ContentAttributes);
                sc.Content.AddRange(s.Content.Select(Clone));
                c = sc;
                break;
            case OpaqueInline o: c = new OpaqueInline(new XElement(o.Xml)); break;
            case BookmarkStart b: c = new BookmarkStart(b.Ns); break;
            case BookmarkEnd b: c = new BookmarkEnd(b.Ns); break;
            default: throw new InvalidOperationException(i.GetType().Name);
        }
        CopyAttrs(i.Attributes, c.Attributes);
        return c;
    }

    public static RunContent Clone(this RunContent rc)
    {
        RunContent c = rc switch
        {
            TextContent t => new TextContent { Value = t.Value, LoadedValue = t.LoadedValue, Deleted = t.Deleted },
            FieldCode f => new FieldCode { Value = f.Value, LoadedValue = f.LoadedValue, Deleted = f.Deleted },
            FieldChar fc => CloneFieldChar(fc),
            TabContent => new TabContent(),
            BreakContent br => new BreakContent { CarriageReturn = br.CarriageReturn },
            OpaqueRunContent o => new OpaqueRunContent(new XElement(o.Xml)),
            DrawingContent d => CloneDrawing(d),
            _ => throw new InvalidOperationException(rc.GetType().Name),
        };
        CopyAttrs(rc.Attributes, c.Attributes);
        return c;
    }

    static DrawingContent CloneDrawing(DrawingContent d)
    {
        var c = new DrawingContent(new XElement(d.Xml));
        foreach (var box in d.TextBoxes)
        {
            var s = new Story();
            s.Blocks.AddRange(box.Blocks.Select(b => b.Clone()));
            c.TextBoxes.Add(s);
        }
        return c;
    }

    static FieldChar CloneFieldChar(FieldChar fc)
    {
        var c = new FieldChar(fc.Ns);
        c.Children.AddRange(fc.Children.Select(x => new XElement(x)));
        return c;
    }
}
