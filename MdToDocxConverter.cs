using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Extensions.Mathematics;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MdTable = Markdig.Extensions.Tables.Table;
using MdTableRow = Markdig.Extensions.Tables.TableRow;
using MdTableCell = Markdig.Extensions.Tables.TableCell;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Md2Docx;

public class MdToDocxConverter
{
    private readonly string _mdDir;
    private readonly Config _cfg;
    private readonly Dictionary<string, StyleSpec> _styleDefs;
    private readonly MemoryStream _docStream;
    private readonly WordprocessingDocument _doc;
    private readonly Body _body;
    private readonly LatexConverter _latex;
    private readonly MarkdownPipeline _pipeline;
    private readonly Action<string>? _warn;

    // Caption pre-scan data: (chapter, idxInChapter, hasH2Before)
    private List<(int chapter, int idx, bool hasH2)> _figSeq = [];
    private List<(int chapter, int idx, bool hasH2)> _tabSeq = [];
    private int _figPop;
    private int _tabPop;

    // Table caption look-back buffer
    private (string text, ContainerInline inline)? _pending;

    public MdToDocxConverter(string mdPath, Config cfg, Action<string>? warn = null)
    {
        _mdDir = Path.GetDirectoryName(Path.GetFullPath(mdPath)) ?? ".";
        _cfg = cfg;
        _warn = warn;
        _styleDefs = StyleDefs.Build(cfg);

        _docStream = new MemoryStream();
        _doc = WordprocessingDocument.Create(_docStream, WordprocessingDocumentType.Document, true);
        var mainPart = _doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        _body = mainPart.Document.Body!;

        // Ensure numbering part exists before injection
        mainPart.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering();

        PageSetup.ClearDefaultBody(_body);
        PageSetup.Setup(_body, cfg.Page);
        StyleInjector.InjectNumbering(_doc, cfg.Numbering, cfg.Fonts, cfg.Spacing);
        StyleInjector.InjectStyles(_doc, _styleDefs);
        StyleInjector.InjectSettings(_doc);

        _latex = new LatexConverter(warn);

        _pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .UseMathematics()
            .Build();
    }

    public void Convert(string mdText)
    {
        mdText = mdText.Replace("\r\n", "\n");
        var document = Markdown.Parse(mdText, _pipeline);

        _latex.Prebuild(CollectFormulas(document));
        (_figSeq, _tabSeq) = PrecomputeCaptionNumbers(document);
        _figPop = 0;
        _tabPop = 0;

        RenderBlocks(document);
        FlushPending();
    }

    public void Save(string outputPath)
    {
        _doc.Dispose();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".");
        File.WriteAllBytes(outputPath, _docStream.ToArray());
    }

    // -----------------------------------------------------------------------
    // Pre-scan: collect all formulas for batch pandoc conversion
    // -----------------------------------------------------------------------

    private static List<(string latex, bool display)> CollectFormulas(MarkdownDocument doc)
    {
        var formulas = new List<(string, bool)>();
        foreach (var block in doc.Descendants<Block>())
        {
            if (block is MathBlock mathBlock)
            {
                formulas.Add((GetBlockText(mathBlock), true));
            }
            else if (block is LeafBlock leaf && leaf.Inline is not null)
            {
                foreach (var inline in leaf.Inline.Descendants<Inline>())
                {
                    if (inline is MathInline math)
                        formulas.Add((math.Content.ToString(), InlineRenderer.IsDisplayMath(math)));
                }
            }
        }
        return formulas;
    }

    // -----------------------------------------------------------------------
    // Pre-scan: compute caption numbers (chapter, idx, hasH2)
    // -----------------------------------------------------------------------

    private static (List<(int, int, bool)> fig, List<(int, int, bool)> tab) PrecomputeCaptionNumbers(
        MarkdownDocument doc)
    {
        int chapter = 0, figIdx = 0, tabIdx = 0;
        bool hasH2 = false;
        var figSeq = new List<(int, int, bool)>();
        var tabSeq = new List<(int, int, bool)>();

        foreach (var block in doc)
        {
            switch (block)
            {
                case HeadingBlock { Level: 1 }:
                    chapter++; figIdx = 0; tabIdx = 0; hasH2 = false;
                    break;
                case HeadingBlock { Level: 2 }:
                    hasH2 = true;
                    break;
                case ParagraphBlock para when IsOnlyImageBlock(para):
                    figIdx++;
                    figSeq.Add((chapter, figIdx, hasH2));
                    break;
                case MdTable:
                    tabIdx++;
                    tabSeq.Add((chapter, tabIdx, hasH2));
                    break;
            }
        }
        return (figSeq, tabSeq);
    }

    private static bool IsOnlyImageBlock(ParagraphBlock para)
        => para.Inline is not null && InlineRenderer.IsOnlyImage(para.Inline);

    // -----------------------------------------------------------------------
    // Block rendering — walk top-level AST
    // -----------------------------------------------------------------------

    private void RenderBlocks(ContainerBlock container)
    {
        foreach (var block in container)
        {
            switch (block)
            {
                case HeadingBlock heading:
                    FlushPending();
                    RenderHeading(heading);
                    break;

                case ParagraphBlock para:
                {
                    var text = InlineRenderer.GetPlainText(para.Inline);
                    if (string.IsNullOrWhiteSpace(text) && !InlineRenderer.HasImage(para.Inline))
                        break;
                    FlushPending();
                    _pending = (text, para.Inline!);
                    break;
                }

                case MdTable table:
                    RenderTable(table);
                    break;

                case MathBlock math:
                    FlushPending();
                    RenderMathBlock(GetBlockText(math));
                    break;

                case CodeBlock code:
                    FlushPending();
                    RenderCodeBlock(GetBlockText(code));
                    break;

                case ListBlock list:
                    FlushPending();
                    RenderList(list, 0);
                    break;

                case QuoteBlock quote:
                    FlushPending();
                    RenderBlockquote(quote);
                    break;

                case ThematicBreakBlock:
                    FlushPending();
                    var hrP = Oxml.MakePara(_body, "body", _styleDefs);
                    hrP.Append(Oxml.BuildRun("────────────────", null));
                    break;
            }
        }
    }

    // -----------------------------------------------------------------------
    // Pending flush (table caption look-back)
    // -----------------------------------------------------------------------

    private void FlushPending(bool asTableCaption = false)
    {
        if (_pending is null) return;
        var (text, inline) = _pending.Value;
        _pending = null;

        if (asTableCaption)
        {
            var p = Oxml.MakePara(_body, "table_caption", _styleDefs);
            InlineRenderer.RenderInto(p, inline, "table_caption", _styleDefs, _cfg.Fonts.Code, _latex);
            return;
        }

        // Display math: single-line $$...$$ that Markdig parses as MathInline
        var displayMath = InlineRenderer.GetDisplayMath(inline);
        if (displayMath is not null)
        {
            RenderMathBlock(displayMath.Content.ToString());
            return;
        }

        if (InlineRenderer.IsOnlyImage(inline))
        {
            var imgLink = InlineRenderer.GetOnlyImage(inline)!;
            var alt = InlineRenderer.GetPlainText(imgLink);
            var src = imgLink.Url ?? "";
            RenderImage(src, alt);
            return;
        }

        Dictionary<string, object?>? ovr = InlineRenderer.IsAllBold(inline)
            ? new() { ["first_line"] = 0, ["first_line_chars"] = 0 }
            : null;
        var para = Oxml.MakePara(_body, "body", _styleDefs, ovr);
        InlineRenderer.RenderInto(para, inline, "body", _styleDefs, _cfg.Fonts.Code, _latex);
    }

    // -----------------------------------------------------------------------
    // Individual block renderers
    // -----------------------------------------------------------------------

    private void RenderHeading(HeadingBlock heading)
    {
        var key = StyleDefs.HeadingKeys.GetValueOrDefault(heading.Level, "h4");
        var p = Oxml.MakePara(_body, key, _styleDefs);
        InlineRenderer.RenderInto(p, heading.Inline, key, _styleDefs, _cfg.Fonts.Code, _latex);
    }

    private void RenderCodeBlock(string content)
    {
        var text = content.TrimEnd('\n');
        foreach (var line in text.Split('\n'))
        {
            var p = Oxml.MakePara(_body, "code_block", _styleDefs);
            p.Append(Oxml.BuildRun(line, null));
        }
    }

    private void RenderList(ListBlock list, int depth)
    {
        bool ordered = list.IsOrdered;
        foreach (var item in list)
        {
            if (item is ListItemBlock listItem)
                RenderListItem(listItem, ordered, depth);
        }
    }

    private void RenderListItem(ListItemBlock item, bool ordered, int depth)
    {
        int numId = ordered ? StyleDefs.OrderedNumId : StyleDefs.BulletNumId;

        foreach (var block in item)
        {
            switch (block)
            {
                case ParagraphBlock para:
                {
                    var ovr = new Dictionary<string, object?>
                    {
                        ["numPr"] = new Dictionary<string, int> { ["ilvl"] = depth, ["numId"] = numId }
                    };
                    var p = Oxml.MakePara(_body, "list_item", _styleDefs, ovr);
                    InlineRenderer.RenderInto(p, para.Inline, "list_item", _styleDefs, _cfg.Fonts.Code, _latex);
                    break;
                }
                case ListBlock nestedList:
                    RenderList(nestedList, depth + 1);
                    break;
                case FencedCodeBlock fenced:
                    RenderCodeBlock(GetBlockText(fenced));
                    break;
                case CodeBlock code:
                    RenderCodeBlock(GetBlockText(code));
                    break;
            }
        }
    }

    private void RenderBlockquote(QuoteBlock quote)
    {
        foreach (var block in quote)
        {
            if (block is ParagraphBlock para)
            {
                var p = Oxml.MakePara(_body, "blockquote", _styleDefs);
                InlineRenderer.RenderInto(p, para.Inline, "blockquote", _styleDefs, _cfg.Fonts.Code, _latex);
            }
        }
    }

    private void RenderMathBlock(string latexContent)
    {
        var ommlXml = _latex.ToOmml(latexContent, display: true);
        var p = Oxml.MakePara(_body, "math_display", _styleDefs);
        if (ommlXml is not null)
        {
            var elem = LatexConverter.ParseOmmlElement(ommlXml);
            if (elem is not null) { p.Append(elem); return; }
        }
        _warn?.Invoke($"display math fallback: {latexContent}");
        var cf = _cfg.Fonts.Code;
        var rpr = new Dictionary<string, object?>
        {
            ["ascii"] = cf, ["hAnsi"] = cf, ["eastAsia"] = cf, ["cs"] = cf,
            ["italic"] = true,
        };
        p.Append(Oxml.BuildRun(latexContent, rpr));
    }

    // -----------------------------------------------------------------------
    // Image + caption
    // -----------------------------------------------------------------------

    private void RenderImage(string src, string alt)
    {
        if (string.IsNullOrEmpty(src)) return;

        var srcPath = Path.IsPathRooted(src) ? src : Path.GetFullPath(Path.Combine(_mdDir, src));

        if (!File.Exists(srcPath))
        {
            _warn?.Invoke($"image not found: {srcPath}");
            var fp = Oxml.MakePara(_body, "image", _styleDefs);
            fp.Append(Oxml.BuildRun($"[图片缺失: {src}]",
                new Dictionary<string, object?> { ["italic"] = true, ["color"] = "999999" }));
        }
        else
        {
            InsertImage(srcPath);
        }

        // Caption: 图 <chapter>-<idx> <alt>
        if (_figPop < _figSeq.Count)
        {
            var (chapter, figIdx, hasH2) = _figSeq[_figPop++];
            var capP = Oxml.MakePara(_body, "image_caption", _styleDefs);
            AppendCaptionPrefix(capP, "图", "Figure", chapter, figIdx, hasH2);
            if (!string.IsNullOrWhiteSpace(alt))
                capP.Append(Oxml.BuildRun(alt.Trim(), null));
        }
    }

    private void InsertImage(string srcPath, double widthCm = 15.0)
    {
        var p = Oxml.MakePara(_body, "image", _styleDefs);

        try
        {
            var mainPart = _doc.MainDocumentPart!;
            var imgPart = mainPart.AddImagePart(ImagePartType.Png);
            using (var fs = File.OpenRead(srcPath))
                imgPart.FeedData(fs);

            var relId = mainPart.GetIdOfPart(imgPart);
            long widthEmu = (long)(widthCm * 360000);
            long heightEmu = (long)(widthEmu * 0.75); // 4:3 default; Word auto-adjusts on open

            var drawing = CreateImageDrawing(relId, widthEmu, heightEmu, Path.GetFileName(srcPath));
            var run = new Run(drawing);
            p.Append(run);
        }
        catch (Exception ex)
        {
            _warn?.Invoke($"add_picture failed for {srcPath}: {ex.Message}");
            p.Append(Oxml.BuildRun($"[图片加载失败: {Path.GetFileName(srcPath)}]",
                new Dictionary<string, object?> { ["italic"] = true, ["color"] = "999999" }));
        }
    }

    private static OpenXmlElement CreateImageDrawing(string relId, long cx, long cy, string name)
    {
        // Build drawing XML directly since the type hierarchy is complex
        var drawingXml = $@"<w:drawing xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
            xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
            xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
            xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture""
            xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
            <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"">
                <wp:extent cx=""{cx}"" cy=""{cy}""/>
                <wp:docPr id=""1"" name=""{System.Security.SecurityElement.Escape(name)}""/>
                <a:graphic>
                    <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                        <pic:pic>
                            <pic:nvPicPr>
                                <pic:cNvPr id=""0"" name=""{System.Security.SecurityElement.Escape(name)}""/>
                                <pic:cNvPicPr/>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed=""{relId}""/>
                                <a:stretch><a:fillRect/></a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x=""0"" y=""0""/>
                                    <a:ext cx=""{cx}"" cy=""{cy}""/>
                                </a:xfrm>
                                <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:inline>
        </w:drawing>";

        var wrapper = new OpenXmlUnknownElement("w", "body",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        { InnerXml = drawingXml };
        return wrapper.FirstChild!.CloneNode(true);
    }

    // -----------------------------------------------------------------------
    // Table + caption
    // -----------------------------------------------------------------------

    private void RenderTable(MdTable table)
    {
        // Look-back: pending paragraph becomes table caption
        var captionPending = _pending;
        _pending = null;

        if (_tabPop < _tabSeq.Count)
        {
            var (chapter, tabIdx, hasH2) = _tabSeq[_tabPop++];
            var capP = Oxml.MakePara(_body, "table_caption", _styleDefs);
            AppendCaptionPrefix(capP, "表", "Table", chapter, tabIdx, hasH2);
            if (captionPending is not null)
                InlineRenderer.RenderInto(capP, captionPending.Value.inline, "table_caption",
                    _styleDefs, _cfg.Fonts.Code, _latex);
        }

        BuildWordTable(table);
    }

    private void BuildWordTable(MdTable mdTable)
    {
        var rows = mdTable.OfType<MdTableRow>().ToList();
        if (rows.Count == 0) return;
        int ncols = rows.Max(r => r.Count);

        var tbl = new W.Table();

        // Table properties
        var tblPr = new TableProperties(
            new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
            new TableLook
            {
                Val = "04A0",
                FirstRow = true, LastRow = false,
                FirstColumn = true, LastColumn = false,
                NoHorizontalBand = false, NoVerticalBand = true,
            },
            BuildTableBorders(),
            new TableCellMarginDefault(
                new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa }
            )
        );
        tbl.Append(tblPr);

        int usable = _cfg.Page.Width - 2 * _cfg.Page.Margin;
        int colW = usable / Math.Max(ncols, 1);

        var grid = new TableGrid();
        for (int i = 0; i < ncols; i++)
            grid.Append(new GridColumn { Width = colW.ToString() });
        tbl.Append(grid);

        foreach (var row in rows)
        {
            bool isHeader = row.IsHeader;
            var tr = new TableRow();
            if (isHeader)
                tr.Append(new TableRowProperties(new TableHeader()));

            var specKey = isHeader ? "cell_header" : "cell_body";
            for (int ci = 0; ci < ncols; ci++)
            {
                var tc = new TableCell();
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Width = colW.ToString(), Type = TableWidthUnitValues.Dxa },
                    BuildCellBorders(),
                    new W.Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                ));

                var cellP = Oxml.MakeStyledP(specKey, _styleDefs);
                if (ci < row.Count && row[ci] is MdTableCell mdCell)
                {
                    foreach (var cellBlock in mdCell)
                    {
                        if (cellBlock is ParagraphBlock para)
                            InlineRenderer.RenderInto(cellP, para.Inline, specKey, _styleDefs,
                                _cfg.Fonts.Code, _latex);
                    }
                }
                tc.Append(cellP);
                tr.Append(tc);
            }
            tbl.Append(tr);
        }

        Oxml.AppendToBody(_body, tbl);
        // Spacer paragraph after table
        var spacer = Oxml.MakeStyledP("body", _styleDefs,
            new Dictionary<string, object?> { ["before"] = 0, ["after"] = 0 });
        Oxml.AppendToBody(_body, spacer);
    }

    private static TableBorders BuildTableBorders()
    {
        static T MakeBorder<T>(T border) where T : BorderType
        {
            border.Val = BorderValues.Single;
            border.Color = "000000";
            border.Space = (UInt32Value)0u;
            border.Size = (UInt32Value)4u;
            return border;
        }
        return new TableBorders(
            MakeBorder(new TopBorder()),
            MakeBorder(new LeftBorder()),
            MakeBorder(new BottomBorder()),
            MakeBorder(new RightBorder()),
            MakeBorder(new InsideHorizontalBorder()),
            MakeBorder(new InsideVerticalBorder())
        );
    }

    private static TableCellBorders BuildCellBorders()
    {
        static T MakeBorder<T>(T border) where T : BorderType
        {
            border.Val = BorderValues.Single;
            border.Color = "000000";
            border.Space = (UInt32Value)0u;
            border.Size = (UInt32Value)4u;
            return border;
        }
        return new TableCellBorders(
            MakeBorder(new TopBorder()),
            MakeBorder(new LeftBorder()),
            MakeBorder(new BottomBorder()),
            MakeBorder(new RightBorder())
        );
    }

    // -----------------------------------------------------------------------
    // Caption prefix: "图 1-1 " / "表 1-1 " with SEQ/STYLEREF fields
    // -----------------------------------------------------------------------

    private static void AppendCaptionPrefix(
        Paragraph capP, string label, string seqName,
        int chapter, int idx, bool hasH2)
    {
        capP.Append(Oxml.BuildRun($"{label} ", null));

        if (chapter == 0 || !hasH2)
        {
            capP.Append(Oxml.BuildRun(chapter.ToString(), null));
        }
        else
        {
            foreach (var r in Oxml.BuildNestedFieldRuns(
                " = INT(", " STYLEREF 2 \\r ", ") ",
                "", chapter.ToString()))
            {
                capP.Append(r);
            }
        }

        capP.Append(Oxml.BuildRun("-", null));
        foreach (var r in Oxml.BuildFieldRuns($" SEQ {seqName} \\s 1 \\* ARABIC ", idx.ToString()))
            capP.Append(r);
        capP.Append(Oxml.BuildRun(" ", null));
    }

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------

    private static string GetBlockText(LeafBlock block)
    {
        if (block.Lines.Count == 0) return "";
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < block.Lines.Count; i++)
        {
            if (i > 0) sb.Append('\n');
            sb.Append(block.Lines.Lines[i].Slice.ToString());
        }
        return sb.ToString();
    }
}
