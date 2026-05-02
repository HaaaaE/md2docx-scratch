using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Md2Docx;

public static class Oxml
{
    private static int GetInt(Dictionary<string, object?> spec, string key, int fallback = 0)
    {
        if (spec.TryGetValue(key, out var v) && v is int i) return i;
        return fallback;
    }

    private static string? GetStr(Dictionary<string, object?> spec, string key)
    {
        if (spec.TryGetValue(key, out var v) && v is string s) return s;
        return null;
    }

    private static bool GetBool(Dictionary<string, object?> spec, string key)
    {
        if (spec.TryGetValue(key, out var v) && v is bool b) return b;
        return false;
    }

    public static bool GetBoolPublic(Dictionary<string, object?> spec, string key) => GetBool(spec, key);

    private static bool Has(Dictionary<string, object?> spec, string key)
        => spec.TryGetValue(key, out var v) && v is not null;

    /// <summary>Build ParagraphProperties from a spec dictionary, following OOXML schema order.</summary>
    public static ParagraphProperties BuildPPr(Dictionary<string, object?> spec)
    {
        var pPr = new ParagraphProperties();

        if (GetBool(spec, "keep_next"))
            pPr.Append(new KeepNext());
        if (GetBool(spec, "keep_lines"))
            pPr.Append(new KeepLines());
        if (GetBool(spec, "page_break_before"))
            pPr.Append(new PageBreakBefore());

        if (spec.TryGetValue("numPr", out var numObj) && numObj is Dictionary<string, int> numSpec)
        {
            var numPr = new NumberingProperties(
                new NumberingLevelReference { Val = numSpec.GetValueOrDefault("ilvl", 0) },
                new NumberingId { Val = numSpec["numId"] }
            );
            pPr.Append(numPr);
        }

        bool hasSpacing = Has(spec, "before") || Has(spec, "after") || Has(spec, "line") || Has(spec, "line_rule");
        if (hasSpacing)
        {
            var spacing = new SpacingBetweenLines();
            if (Has(spec, "before")) spacing.Before = GetInt(spec, "before").ToString();
            if (Has(spec, "after")) spacing.After = GetInt(spec, "after").ToString();
            if (Has(spec, "line")) spacing.Line = GetInt(spec, "line").ToString();
            if (Has(spec, "line_rule"))
            {
                var rule = GetStr(spec, "line_rule");
                spacing.LineRule = rule switch
                {
                    "auto" => LineSpacingRuleValues.Auto,
                    "exact" => LineSpacingRuleValues.Exact,
                    "atLeast" => LineSpacingRuleValues.AtLeast,
                    _ => LineSpacingRuleValues.Auto,
                };
            }
            pPr.Append(spacing);
        }

        bool hasInd = Has(spec, "first_line") || Has(spec, "first_line_chars")
                    || Has(spec, "left") || Has(spec, "right") || Has(spec, "hanging");
        if (hasInd)
        {
            var ind = new Indentation();
            if (Has(spec, "left")) ind.Left = GetInt(spec, "left").ToString();
            if (Has(spec, "right")) ind.Right = GetInt(spec, "right").ToString();
            if (Has(spec, "hanging"))
            {
                ind.Hanging = GetInt(spec, "hanging").ToString();
            }
            else
            {
                if (Has(spec, "first_line")) ind.FirstLine = GetInt(spec, "first_line").ToString();
                if (Has(spec, "first_line_chars")) ind.FirstLineChars = GetInt(spec, "first_line_chars");
            }
            pPr.Append(ind);
        }

        var align = GetStr(spec, "align");
        if (align is not null)
        {
            var jc = new Justification
            {
                Val = align switch
                {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "both" => JustificationValues.Both,
                    _ => JustificationValues.Left,
                }
            };
            pPr.Append(jc);
        }

        if (Has(spec, "outline"))
        {
            pPr.Append(new OutlineLevel { Val = GetInt(spec, "outline") });
        }

        return pPr;
    }

    /// <summary>Build RunProperties from a spec dictionary. Returns null if spec is null/empty.</summary>
    public static RunProperties? BuildRPr(Dictionary<string, object?>? spec)
    {
        if (spec is null || spec.Count == 0) return null;
        var rPr = new RunProperties();

        bool hasFonts = Has(spec, "ascii") || Has(spec, "hAnsi") || Has(spec, "eastAsia") || Has(spec, "cs");
        if (hasFonts)
        {
            var fonts = new RunFonts();
            if (Has(spec, "ascii")) fonts.Ascii = GetStr(spec, "ascii");
            if (Has(spec, "hAnsi")) fonts.HighAnsi = GetStr(spec, "hAnsi");
            else if (Has(spec, "ascii")) fonts.HighAnsi = GetStr(spec, "ascii");
            if (Has(spec, "eastAsia")) fonts.EastAsia = GetStr(spec, "eastAsia");
            if (Has(spec, "cs")) fonts.ComplexScript = GetStr(spec, "cs");
            rPr.Append(fonts);
        }

        if (GetBool(spec, "bold"))
        {
            rPr.Append(new Bold());
            rPr.Append(new BoldComplexScript());
        }
        if (GetBool(spec, "italic"))
        {
            rPr.Append(new Italic());
            rPr.Append(new ItalicComplexScript());
        }
        if (GetBool(spec, "strike"))
            rPr.Append(new Strike());

        if (Has(spec, "size"))
        {
            var sz = GetInt(spec, "size").ToString();
            rPr.Append(new FontSize { Val = sz });
            rPr.Append(new FontSizeComplexScript { Val = sz });
        }

        if (Has(spec, "color"))
            rPr.Append(new W.Color { Val = GetStr(spec, "color") });

        if (Has(spec, "shd"))
        {
            rPr.Append(new W.Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = GetStr(spec, "shd"),
            });
        }

        if (GetBool(spec, "underline"))
            rPr.Append(new Underline { Val = UnderlineValues.Single });

        return rPr;
    }

    /// <summary>Build a Run with text and optional rPr spec. Splits on newlines using Break elements.</summary>
    public static Run BuildRun(string text, Dictionary<string, object?>? rprSpec)
    {
        var run = new Run();
        var rPr = BuildRPr(rprSpec);
        if (rPr is not null) run.Append(rPr);

        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            if (i > 0) run.Append(new Break());
            if (lines[i].Length > 0)
                run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
        }
        return run;
    }

    /// <summary>Build the 5-run field sequence: begin / instrText / separate / result / end.</summary>
    public static List<Run> BuildFieldRuns(string instrText, string resultText, bool hidden = false)
    {
        Run MakeRun(OpenXmlElement child)
        {
            var r = new Run();
            if (hidden)
            {
                var rPr = new RunProperties();
                rPr.Append(new Vanish());
                r.Append(rPr);
            }
            r.Append(child);
            return r;
        }

        return
        [
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            MakeRun(new FieldCode(instrText) { Space = SpaceProcessingModeValues.Preserve }),
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            MakeRun(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve }),
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.End }),
        ];
    }

    /// <summary>Build the 11-run nested field sequence: { outerPrefix { innerInstr } outerSuffix }.</summary>
    public static List<Run> BuildNestedFieldRuns(
        string outerPrefix, string innerInstr, string outerSuffix,
        string innerResult, string outerResult)
    {
        static Run MakeRun(OpenXmlElement child)
        {
            var r = new Run();
            r.Append(child);
            return r;
        }

        return
        [
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            MakeRun(new FieldCode(outerPrefix) { Space = SpaceProcessingModeValues.Preserve }),
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            MakeRun(new FieldCode(innerInstr) { Space = SpaceProcessingModeValues.Preserve }),
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            MakeRun(new Text(innerResult) { Space = SpaceProcessingModeValues.Preserve }),
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.End }),
            MakeRun(new FieldCode(outerSuffix) { Space = SpaceProcessingModeValues.Preserve }),
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            MakeRun(new Text(outerResult) { Space = SpaceProcessingModeValues.Preserve }),
            MakeRun(new FieldChar { FieldCharType = FieldCharValues.End }),
        ];
    }

    /// <summary>Insert element before the final SectionProperties in body.</summary>
    public static void AppendToBody(Body body, OpenXmlElement element)
    {
        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectPr is not null)
            body.InsertBefore(element, sectPr);
        else
            body.Append(element);
    }

    /// <summary>Merge two rPr spec dictionaries (override wins).</summary>
    public static Dictionary<string, object?>? MergeRPr(
        Dictionary<string, object?>? baseSpec, Dictionary<string, object?>? overrideSpec)
    {
        if (baseSpec is null && overrideSpec is null) return null;
        var result = new Dictionary<string, object?>();
        if (baseSpec is not null)
            foreach (var kv in baseSpec) result[kv.Key] = kv.Value;
        if (overrideSpec is not null)
            foreach (var kv in overrideSpec) result[kv.Key] = kv.Value;
        return result;
    }

    /// <summary>Create a paragraph with pStyle set, optionally with direct pPr overrides.</summary>
    public static Paragraph MakeStyledP(string specKey, Dictionary<string, StyleSpec> styleDefs,
        Dictionary<string, object?>? pPrOverride = null)
    {
        var styleId = styleDefs[specKey].StyleId;
        var p = new Paragraph();
        var pPr = new ParagraphProperties(
            new ParagraphStyleId { Val = styleId }
        );
        if (pPrOverride is not null)
        {
            var overridePPr = BuildPPr(pPrOverride);
            foreach (var child in overridePPr.ChildElements.ToList())
            {
                child.Remove();
                pPr.Append(child);
            }
        }
        p.Append(pPr);
        return p;
    }

    /// <summary>Create a styled paragraph and append it to body.</summary>
    public static Paragraph MakePara(Body body, string specKey, Dictionary<string, StyleSpec> styleDefs,
        Dictionary<string, object?>? pPrOverride = null)
    {
        var p = MakeStyledP(specKey, styleDefs, pPrOverride);
        AppendToBody(body, p);
        return p;
    }
}
