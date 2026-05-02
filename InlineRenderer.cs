using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Extensions.Mathematics;
using Markdig.Syntax.Inlines;

namespace Md2Docx;

public static class InlineRenderer
{
    /// <summary>
    /// Render inline content into a paragraph. Walks the Markdig inline AST recursively,
    /// emitting Word runs with appropriate formatting.
    /// </summary>
    public static void RenderInto(
        Paragraph p,
        ContainerInline? inlines,
        string parentStyleKey,
        Dictionary<string, StyleSpec> styleDefs,
        string codeFont,
        LatexConverter latex)
    {
        if (inlines is null) return;

        var styleDef = styleDefs.GetValueOrDefault(parentStyleKey);
        var styleRPr = styleDef?.RPr;
        bool styleBold = styleRPr is not null && Oxml.GetBoolPublic(styleRPr, "bold");
        bool styleItalic = styleRPr is not null && Oxml.GetBoolPublic(styleRPr, "italic");

        foreach (var inline in inlines)
            WalkInline(inline, p, styleBold, styleItalic, false, styleBold, styleItalic, codeFont, latex);
    }

    private static void WalkInline(
        Inline inline, Paragraph p,
        bool bold, bool italic, bool strike,
        bool styleBold, bool styleItalic,
        string codeFont, LatexConverter latex)
    {
        switch (inline)
        {
            case LiteralInline lit:
            {
                var text = lit.Content.ToString();
                if (string.IsNullOrEmpty(text)) break;
                var rpr = BuildBaseRPr(bold, italic, strike, styleBold, styleItalic);
                p.Append(Oxml.BuildRun(text, rpr));
                break;
            }

            case EmphasisInline emp:
            {
                // DelimiterCount==2 → bold (**), DelimiterCount==1 → italic (*),
                // DelimiterChar=='~' → strikethrough (~~)
                bool newBold = bold;
                bool newItalic = italic;
                bool newStrike = strike;
                if (emp.DelimiterChar is '*' or '_')
                {
                    if (emp.DelimiterCount == 2) newBold = true;
                    else newItalic = true;
                }
                else if (emp.DelimiterChar == '~' && emp.DelimiterCount == 2)
                {
                    newStrike = true;
                }

                foreach (var child in emp)
                    WalkInline(child, p, newBold, newItalic, newStrike, styleBold, styleItalic, codeFont, latex);
                break;
            }

            case CodeInline code:
            {
                var rpr = BuildCodeRPr(codeFont);
                p.Append(Oxml.BuildRun(code.Content ?? "", rpr));
                break;
            }

            case MathInline math:
            {
                var content = math.Content.ToString();
                var ommlXml = latex.ToOmml(content, display: false);
                if (ommlXml is not null)
                {
                    var elem = LatexConverter.ParseOmmlElement(ommlXml);
                    if (elem is not null) { p.Append(elem); break; }
                }
                // Fallback: italic text
                var rpr = BuildBaseRPr(bold, italic, strike, styleBold, styleItalic);
                rpr ??= new Dictionary<string, object?>();
                if (!styleItalic) rpr["italic"] = true;
                p.Append(Oxml.BuildRun(content, rpr));
                break;
            }

            case LineBreakInline lb:
            {
                if (lb.IsHard)
                {
                    p.Append(Oxml.BuildRun("\n", BuildBaseRPr(bold, italic, strike, styleBold, styleItalic)));
                }
                else
                {
                    p.Append(Oxml.BuildRun(" ", BuildBaseRPr(bold, italic, strike, styleBold, styleItalic)));
                }
                break;
            }

            case LinkInline link when !link.IsImage:
            {
                foreach (var child in link)
                    WalkInline(child, p, bold, italic, strike, styleBold, styleItalic, codeFont, latex);
                break;
            }

            case ContainerInline container:
            {
                foreach (var child in container)
                    WalkInline(child, p, bold, italic, strike, styleBold, styleItalic, codeFont, latex);
                break;
            }

            case HtmlInline html:
            {
                p.Append(Oxml.BuildRun(html.Tag ?? "", BuildBaseRPr(bold, italic, strike, styleBold, styleItalic)));
                break;
            }
        }
    }

    /// <summary>
    /// Build rPr for a plain text run. Only emits bold/italic/strike if the inline state
    /// differs from the parent style (avoiding OOXML toggle XOR issues).
    /// </summary>
    private static Dictionary<string, object?>? BuildBaseRPr(
        bool bold, bool italic, bool strike,
        bool styleBold, bool styleItalic)
    {
        Dictionary<string, object?>? rpr = null;
        if (bold && !styleBold) (rpr ??= new())["bold"] = true;
        if (italic && !styleItalic) (rpr ??= new())["italic"] = true;
        if (strike) (rpr ??= new())["strike"] = true;
        return rpr;
    }

    private static Dictionary<string, object?> BuildCodeRPr(string codeFont)
    {
        return new Dictionary<string, object?>
        {
            ["ascii"] = codeFont, ["hAnsi"] = codeFont,
            ["eastAsia"] = codeFont, ["cs"] = codeFont,
            ["shd"] = "F2F2F2",
        };
    }

    /// <summary>Markdig parses single-line $$...$$ as MathInline with two delimiters.</summary>
    public static bool IsDisplayMath(MathInline math) => math.DelimiterCount == 2;

    /// <summary>
    /// Check if the paragraph is a display math paragraph:
    /// a single MathInline with $$ delimiters (Markdig treats single-line $$...$$ as inline).
    /// </summary>
    public static MathInline? GetDisplayMath(ContainerInline? inlines)
    {
        if (inlines is null) return null;
        MathInline? found = null;
        foreach (var inline in inlines)
        {
            if (inline is LiteralInline lit && string.IsNullOrWhiteSpace(lit.Content.ToString()))
                continue;
            if (inline is MathInline math && found is null)
                found = math;
            else
                return null;
        }
        if (found is not null && IsDisplayMath(found))
            return found;
        return null;
    }

    /// <summary>Check if a ParagraphBlock's inline content is a single image.</summary>
    public static bool IsOnlyImage(ContainerInline? inlines)
    {
        if (inlines is null) return false;
        LinkInline? img = null;
        foreach (var inline in inlines)
        {
            if (inline is LiteralInline lit && string.IsNullOrWhiteSpace(lit.Content.ToString()))
                continue;
            if (inline is LinkInline { IsImage: true } link && img is null)
                img = link;
            else
                return false;
        }
        return img is not null;
    }

    /// <summary>Check if any inline is an image.</summary>
    public static bool HasImage(ContainerInline? inlines)
    {
        if (inlines is null) return false;
        foreach (var inline in inlines)
            if (inline is LinkInline { IsImage: true }) return true;
        return false;
    }

    /// <summary>Get the single image LinkInline from inline content.</summary>
    public static LinkInline? GetOnlyImage(ContainerInline? inlines)
    {
        if (inlines is null) return null;
        foreach (var inline in inlines)
            if (inline is LinkInline { IsImage: true } link) return link;
        return null;
    }

    /// <summary>Get plain text from inline content.</summary>
    public static string GetPlainText(ContainerInline? inlines)
    {
        if (inlines is null) return "";
        var parts = new List<string>();
        CollectText(inlines, parts);
        return string.Join("", parts);
    }

    private static void CollectText(Inline inline, List<string> parts)
    {
        switch (inline)
        {
            case LiteralInline lit:
                parts.Add(lit.Content.ToString());
                break;
            case CodeInline code:
                parts.Add(code.Content ?? "");
                break;
            case LineBreakInline lb:
                parts.Add(lb.IsHard ? "\n" : " ");
                break;
            case MathInline math:
                parts.Add(math.Content.ToString());
                break;
            case ContainerInline container:
                foreach (var child in container) CollectText(child, parts);
                break;
        }
    }

    /// <summary>
    /// Check if the entire paragraph is fully bold text (all LiteralInline inside EmphasisInline with count==2).
    /// Used to suppress first-line indent for emphasis paragraphs.
    /// </summary>
    public static bool IsAllBold(ContainerInline? inlines)
    {
        if (inlines is null) return false;
        bool hasBoldText = false;
        return CheckAllBold(inlines, false, ref hasBoldText) && hasBoldText;
    }

    private static bool CheckAllBold(Inline inline, bool inBold, ref bool hasBoldText)
    {
        switch (inline)
        {
            case LiteralInline lit:
                if (string.IsNullOrWhiteSpace(lit.Content.ToString())) return true;
                if (!inBold) return false;
                hasBoldText = true;
                return true;

            case EmphasisInline emp when emp.DelimiterChar is '*' or '_' && emp.DelimiterCount == 2:
                foreach (var child in emp)
                    if (!CheckAllBold(child, true, ref hasBoldText)) return false;
                return true;

            case EmphasisInline emp:
                foreach (var child in emp)
                    if (!CheckAllBold(child, inBold, ref hasBoldText)) return false;
                return true;

            case LineBreakInline:
                return true;

            case LinkInline link when !link.IsImage:
                foreach (var child in link)
                    if (!CheckAllBold(child, inBold, ref hasBoldText)) return false;
                return true;

            case ContainerInline container:
                foreach (var child in container)
                    if (!CheckAllBold(child, inBold, ref hasBoldText)) return false;
                return true;

            default:
                return false;
        }
    }
}
