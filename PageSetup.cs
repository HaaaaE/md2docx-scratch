using DocumentFormat.OpenXml.Wordprocessing;

namespace Md2Docx;

public static class PageSetup
{
    public static void Setup(Body body, PageConfig pg)
    {
        var oldSect = body.Elements<SectionProperties>().FirstOrDefault();
        oldSect?.Remove();

        var sectPr = new SectionProperties(
            new PageSize
            {
                Width = (uint)pg.Width,
                Height = (uint)pg.Height,
            },
            new PageMargin
            {
                Top = pg.Margin,
                Right = (uint)pg.Margin,
                Bottom = pg.Margin,
                Left = (uint)pg.Margin,
                Header = (uint)pg.HeaderDist,
                Footer = (uint)pg.FooterDist,
                Gutter = (uint)pg.Gutter,
            },
            new Columns { Space = "720" },
            new DocGrid { LinePitch = 312 }
        );
        body.Append(sectPr);
    }

    public static void ClearDefaultBody(Body body)
    {
        foreach (var p in body.Elements<Paragraph>().ToList())
            p.Remove();
        foreach (var t in body.Elements<Table>().ToList())
            t.Remove();
    }
}
