namespace Md2Docx;

public record StyleSpec
{
    public required string StyleId { get; init; }
    public required string Name { get; init; }
    public string? BasedOn { get; init; }
    public string? Next { get; init; }
    public string? Link { get; init; }
    public bool Custom { get; init; }
    public bool Default { get; init; }
    public int UiPriority { get; init; }
    public bool QFormat { get; init; }
    public Dictionary<string, object?>? PPr { get; init; }
    public Dictionary<string, object?>? RPr { get; init; }
}

public static class StyleDefs
{
    public const int H1AbstractNumId = 100;
    public const int H1NumId = 100;
    public const int BulletAbstractNumId = 101;
    public const int BulletNumId = 101;
    public const int OrderedAbstractNumId = 102;
    public const int OrderedNumId = 102;

    public static readonly Dictionary<int, string> HeadingKeys = new()
    {
        [1] = "h1", [2] = "h2", [3] = "h3", [4] = "h4", [5] = "h4", [6] = "h4"
    };

    public static Dictionary<string, StyleSpec> Build(Config cfg)
    {
        var f = cfg.Fonts;
        var s = cfg.Sizes;
        var c = cfg.Colors;
        var sp = cfg.Spacing;

        return new Dictionary<string, StyleSpec>
        {
            ["body"] = new()
            {
                StyleId = "Normal", Name = "Normal", Default = true,
                QFormat = true, UiPriority = 0,
                PPr = new()
                {
                    ["before"] = sp.BodyBefore, ["after"] = sp.BodyAfter,
                    ["line"] = sp.BodyLine, ["line_rule"] = sp.BodyLineRule,
                    ["first_line"] = sp.BodyFirstLine, ["first_line_chars"] = sp.BodyFirstLineChars,
                },
                RPr = new()
                {
                    ["ascii"] = f.Latin, ["hAnsi"] = f.Latin, ["cs"] = f.Latin, ["eastAsia"] = f.CjkBody,
                    ["size"] = s.Body,
                },
            },
            ["h1"] = new()
            {
                StyleId = "Heading1", Name = "heading 1",
                BasedOn = "Normal", Next = "Normal", Link = "Heading1Char",
                UiPriority = 9, QFormat = true,
                PPr = new()
                {
                    ["align"] = "center",
                    ["before"] = sp.H1Before, ["after"] = sp.H1After,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                    ["keep_next"] = true, ["keep_lines"] = true,
                    ["outline"] = 0,
                    ["numPr"] = new Dictionary<string, int> { ["ilvl"] = 0, ["numId"] = H1NumId },
                },
                RPr = new()
                {
                    ["ascii"] = f.Latin, ["hAnsi"] = f.Latin, ["cs"] = f.Latin, ["eastAsia"] = f.CjkH1,
                    ["size"] = s.H1, ["color"] = c.H1, ["bold"] = true,
                },
            },
            ["h2"] = new()
            {
                StyleId = "Heading2", Name = "heading 2",
                BasedOn = "Normal", Next = "Normal", Link = "Heading2Char",
                UiPriority = 9, QFormat = true,
                PPr = new()
                {
                    ["before"] = sp.H2Before, ["after"] = sp.H2After,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                    ["keep_next"] = true, ["keep_lines"] = true,
                    ["outline"] = 1,
                    ["numPr"] = new Dictionary<string, int> { ["ilvl"] = 1, ["numId"] = H1NumId },
                },
                RPr = new()
                {
                    ["ascii"] = f.Latin, ["hAnsi"] = f.Latin, ["cs"] = f.Latin, ["eastAsia"] = f.CjkBody,
                    ["size"] = s.Body, ["bold"] = true,
                },
            },
            ["h3"] = new()
            {
                StyleId = "Heading3", Name = "heading 3",
                BasedOn = "Normal", Next = "Normal", Link = "Heading3Char",
                UiPriority = 9, QFormat = true,
                PPr = new()
                {
                    ["before"] = sp.H3Before, ["after"] = sp.H3After,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                    ["keep_next"] = true, ["keep_lines"] = true,
                    ["outline"] = 2,
                    ["numPr"] = new Dictionary<string, int> { ["ilvl"] = 2, ["numId"] = H1NumId },
                },
                RPr = new()
                {
                    ["ascii"] = f.Latin, ["hAnsi"] = f.Latin, ["cs"] = f.Latin, ["eastAsia"] = f.CjkBody,
                    ["size"] = s.Body, ["bold"] = true,
                },
            },
            ["h4"] = new()
            {
                StyleId = "Heading4", Name = "heading 4",
                BasedOn = "Normal", Next = "Normal", Link = "Heading4Char",
                UiPriority = 9, QFormat = true,
                PPr = new()
                {
                    ["before"] = sp.H4Before, ["after"] = sp.H4After,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                    ["keep_next"] = true, ["keep_lines"] = true,
                    ["outline"] = 3,
                    ["numPr"] = new Dictionary<string, int> { ["ilvl"] = 3, ["numId"] = H1NumId },
                },
                RPr = new()
                {
                    ["ascii"] = f.Latin, ["hAnsi"] = f.Latin, ["cs"] = f.Latin, ["eastAsia"] = f.CjkBody,
                    ["size"] = s.Body, ["color"] = c.H4, ["italic"] = true,
                },
            },
            ["list_item"] = new()
            {
                StyleId = "ListParagraph", Name = "List Paragraph",
                BasedOn = "Normal", Next = "ListParagraph",
                UiPriority = 34, QFormat = true,
                PPr = new()
                {
                    ["before"] = sp.ListBefore, ["after"] = sp.ListAfter,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                },
                RPr = null,
            },
            ["blockquote"] = new()
            {
                StyleId = "Quote", Name = "Quote",
                BasedOn = "Normal", Next = "Normal",
                UiPriority = 29, QFormat = true,
                PPr = new()
                {
                    ["before"] = sp.BlockquoteBefore, ["after"] = sp.BlockquoteAfter,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                    ["left"] = sp.BlockquoteLeft,
                },
                RPr = new() { ["italic"] = true },
            },
            ["image"] = new()
            {
                StyleId = "Image", Name = "Image",
                BasedOn = "Normal", Next = "Normal",
                Custom = true, UiPriority = 11, QFormat = true,
                PPr = new()
                {
                    ["align"] = "center",
                    ["before"] = 0, ["after"] = 60,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                    ["keep_next"] = true, ["keep_lines"] = true,
                },
                RPr = null,
            },
            ["image_caption"] = new()
            {
                StyleId = "ImageCaption", Name = "Image Caption",
                BasedOn = "Normal", Next = "Normal",
                Custom = true, UiPriority = 10, QFormat = true,
                PPr = new()
                {
                    ["align"] = "center",
                    ["before"] = 0, ["after"] = 120,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                },
                RPr = null,
            },
            ["table_caption"] = new()
            {
                StyleId = "TableCaption", Name = "Table Caption",
                BasedOn = "Normal", Next = "Normal",
                Custom = true, UiPriority = 10, QFormat = true,
                PPr = new()
                {
                    ["align"] = "center",
                    ["before"] = 60, ["after"] = 60,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                    ["keep_next"] = true, ["keep_lines"] = true,
                },
                RPr = null,
            },
            ["cell_header"] = new()
            {
                StyleId = "TableHeaderCell", Name = "Table Header Cell",
                BasedOn = "Normal", Next = "Normal",
                Custom = true, UiPriority = 11, QFormat = true,
                PPr = new()
                {
                    ["align"] = "center",
                    ["before"] = 60, ["after"] = 60,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                },
                RPr = new() { ["color"] = "000000", ["bold"] = true },
            },
            ["cell_body"] = new()
            {
                StyleId = "TableBodyCell", Name = "Table Body Cell",
                BasedOn = "Normal", Next = "Normal",
                Custom = true, UiPriority = 11, QFormat = true,
                PPr = new()
                {
                    ["align"] = "left",
                    ["before"] = 60, ["after"] = 60,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                },
                RPr = new() { ["color"] = "000000" },
            },
            ["math_display"] = new()
            {
                StyleId = "MathDisplay", Name = "Math Display",
                BasedOn = "Normal", Next = "Normal",
                Custom = true, UiPriority = 11, QFormat = true,
                PPr = new()
                {
                    ["align"] = "center",
                    ["before"] = 60, ["after"] = 120,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                },
                RPr = null,
            },
            ["code_block"] = new()
            {
                StyleId = "CodeBlock", Name = "Code Block",
                BasedOn = "Normal", Next = "Normal",
                Custom = true, UiPriority = 11, QFormat = true,
                PPr = new()
                {
                    ["before"] = sp.CodeBlockBefore, ["after"] = sp.CodeBlockAfter,
                    ["first_line"] = 0, ["first_line_chars"] = 0,
                    ["left"] = sp.CodeBlockLeft, ["right"] = sp.CodeBlockRight,
                },
                RPr = new()
                {
                    ["ascii"] = f.Code, ["hAnsi"] = f.Code, ["cs"] = f.Code, ["eastAsia"] = f.Code,
                    ["size"] = s.Code, ["shd"] = "F2F2F2",
                },
            },
        };
    }
}
