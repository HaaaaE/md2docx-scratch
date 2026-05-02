using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Md2Docx;

public static class StyleInjector
{
    // -----------------------------------------------------------------------
    // InjectStyles
    // -----------------------------------------------------------------------

    public static void InjectStyles(WordprocessingDocument doc, Dictionary<string, StyleSpec> defs)
    {
        var stylesPart = doc.MainDocumentPart!.StyleDefinitionsPart
            ?? doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();
        var stylesRoot = stylesPart.Styles;

        var existing = new Dictionary<string, W.Style>();
        foreach (var st in stylesRoot.Elements<W.Style>())
        {
            var sid = st.StyleId?.Value;
            if (sid is not null) existing[sid] = st;
        }

        foreach (var (_, sdef) in defs)
        {
            var newStyle = BuildStyleElement(sdef);
            if (existing.TryGetValue(sdef.StyleId, out var old))
            {
                stylesRoot.InsertBefore(newStyle, old);
                old.Remove();
            }
            else
            {
                stylesRoot.Append(newStyle);
            }
        }
    }

    private static W.Style BuildStyleElement(StyleSpec sdef)
    {
        var style = new W.Style
        {
            Type = StyleValues.Paragraph,
            StyleId = sdef.StyleId,
        };
        if (sdef.Default) style.Default = true;
        if (sdef.Custom) style.CustomStyle = true;

        style.Append(new StyleName { Val = sdef.Name });

        if (sdef.BasedOn is not null)
            style.Append(new BasedOn { Val = sdef.BasedOn });
        if (sdef.Next is not null)
            style.Append(new NextParagraphStyle { Val = sdef.Next });
        if (sdef.Link is not null)
            style.Append(new LinkedStyle { Val = sdef.Link });

        style.Append(new UIPriority { Val = sdef.UiPriority });

        if (sdef.QFormat)
            style.Append(new PrimaryStyle());

        if (sdef.PPr is not null)
        {
            var pPr = Oxml.BuildPPr(sdef.PPr);
            // Wrap as StyleParagraphProperties
            var stylePPr = new StyleParagraphProperties();
            foreach (var child in pPr.ChildElements.ToList())
            {
                child.Remove();
                stylePPr.Append(child);
            }
            style.Append(stylePPr);
        }

        if (sdef.RPr is not null)
        {
            var rPr = Oxml.BuildRPr(sdef.RPr);
            if (rPr is not null)
            {
                var styleRPr = new StyleRunProperties();
                foreach (var child in rPr.ChildElements.ToList())
                {
                    child.Remove();
                    styleRPr.Append(child);
                }
                style.Append(styleRPr);
            }
        }

        return style;
    }

    // -----------------------------------------------------------------------
    // InjectNumbering — heading multilevel + bullet + ordered lists
    // -----------------------------------------------------------------------

    /// <param name="spacingCfg">
    /// List text indent: <c>w:left</c> = <see cref="SpacingConfig.BodyFirstLine"/> + ilvl×480 (aligned with normal body).
    /// <see cref="SpacingConfig.ListMarkerTextGap"/> is added only to <c>w:hanging</c>, widening the marker tab field so bullets and numbers sit farther from the text without shifting text past body indent.
    /// </param>
    public static void InjectNumbering(
        WordprocessingDocument doc,
        NumberingConfig numCfg,
        FontsConfig fontsCfg,
        SpacingConfig spacingCfg)
    {
        var numPart = doc.MainDocumentPart!.NumberingDefinitionsPart
            ?? doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
        numPart.Numbering ??= new Numbering();
        var root = numPart.Numbering;

        // Remove existing entries with our IDs (idempotent)
        RemoveById<AbstractNum>(root, e => e.AbstractNumberId?.Value,
            StyleDefs.H1AbstractNumId, StyleDefs.BulletAbstractNumId, StyleDefs.OrderedAbstractNumId);
        RemoveById<NumberingInstance>(root, e => e.NumberID?.Value,
            StyleDefs.H1NumId, StyleDefs.BulletNumId, StyleDefs.OrderedNumId);

        // ── Heading multilevel ──
        var headingAbs = BuildHeadingAbstractNum(numCfg);
        var headingNum = new NumberingInstance(
            new AbstractNumId { Val = StyleDefs.H1AbstractNumId }
        ) { NumberID = StyleDefs.H1NumId };

        // ── Bullet list ──
        var bulletAbs = BuildBulletAbstractNum(fontsCfg, spacingCfg);
        var bulletNum = new NumberingInstance(
            new AbstractNumId { Val = StyleDefs.BulletAbstractNumId }
        ) { NumberID = StyleDefs.BulletNumId };

        // ── Ordered list ──
        var orderedAbs = BuildOrderedAbstractNum(spacingCfg);
        var orderedNum = new NumberingInstance(
            new AbstractNumId { Val = StyleDefs.OrderedAbstractNumId }
        ) { NumberID = StyleDefs.OrderedNumId };

        // Insert abstractNums after existing ones, then nums at the end
        var lastAbs = root.Elements<AbstractNum>().LastOrDefault();
        if (lastAbs is not null)
        {
            root.InsertAfter(orderedAbs, lastAbs);
            root.InsertAfter(bulletAbs, lastAbs);
            root.InsertAfter(headingAbs, lastAbs);
        }
        else
        {
            root.Append(headingAbs);
            root.Append(bulletAbs);
            root.Append(orderedAbs);
        }
        root.Append(headingNum);
        root.Append(bulletNum);
        root.Append(orderedNum);
    }

    private static AbstractNum BuildHeadingAbstractNum(NumberingConfig cfg)
    {
        var abs = new AbstractNum { AbstractNumberId = StyleDefs.H1AbstractNumId };
        abs.Append(new MultiLevelType { Val = MultiLevelValues.Multilevel });

        var levels = new (int ilvl, string pStyle, string numFmt, string lvlText, string suff, bool isLgl)[]
        {
            (0, "Heading1", cfg.H1NumFmt, cfg.H1LvlText, cfg.H1Suff, false),
            (1, "Heading2", "decimal", "%1.%2", "space", true),
            (2, "Heading3", "decimal", "%1.%2.%3", "space", true),
            (3, "Heading4", "decimal", "%1.%2.%3.%4", "space", true),
        };

        foreach (var (ilvl, pStyle, numFmt, lvlText, suff, isLgl) in levels)
        {
            var lvl = new Level { LevelIndex = ilvl };
            lvl.Append(new StartNumberingValue { Val = 1 });
            lvl.Append(new NumberingFormat { Val = ParseNumberFormat(numFmt) });
            lvl.Append(new ParagraphStyleIdInLevel { Val = pStyle });
            if (isLgl) lvl.Append(new IsLegalNumberingStyle());
            lvl.Append(new LevelSuffix { Val = ParseSuffix(suff) });
            lvl.Append(new LevelText { Val = lvlText });
            lvl.Append(new LevelJustification { Val = LevelJustificationValues.Left });
            lvl.Append(new PreviousParagraphProperties(
                new Indentation { Left = "0", FirstLine = "0" }
            ));
            abs.Append(lvl);
        }
        return abs;
    }

    private static AbstractNum BuildBulletAbstractNum(FontsConfig fontsCfg, SpacingConfig spacingCfg)
    {
        var abs = new AbstractNum { AbstractNumberId = StyleDefs.BulletAbstractNumId };
        abs.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        const int hangingBase = 240;
        int gap = spacingCfg.ListMarkerTextGap;
        int hangingVal = hangingBase + gap;
        int bodyIndent = spacingCfg.BodyFirstLine;
        for (int ilvl = 0; ilvl < 3; ilvl++)
        {
            int leftVal = bodyIndent + ilvl * 480;
            var lvl = new Level { LevelIndex = ilvl };
            lvl.Append(new StartNumberingValue { Val = 1 });
            lvl.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
            lvl.Append(new LevelText { Val = "\u00B7" });
            lvl.Append(new LevelJustification { Val = LevelJustificationValues.Left });
            lvl.Append(new LevelSuffix { Val = LevelSuffixValues.Tab });
            lvl.Append(new PreviousParagraphProperties(
                new Tabs(new TabStop
                {
                    Val = TabStopValues.Number,
                    Position = leftVal,
                }),
                new Indentation
                {
                    Left = leftVal.ToString(),
                    Hanging = hangingVal.ToString(),
                }
            ));
            var bulletFont = fontsCfg.CjkBody;
            lvl.Append(new NumberingSymbolRunProperties(
                new RunFonts
                {
                    Ascii = bulletFont, HighAnsi = bulletFont,
                    EastAsia = bulletFont, Hint = FontTypeHintValues.EastAsia,
                }
            ));
            abs.Append(lvl);
        }
        return abs;
    }

    private static AbstractNum BuildOrderedAbstractNum(SpacingConfig spacingCfg)
    {
        var abs = new AbstractNum { AbstractNumberId = StyleDefs.OrderedAbstractNumId };
        abs.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        var fmts = new (string fmt, string txt)[]
        {
            ("decimal", "%1."),
            ("lowerLetter", "%2)"),
            ("lowerRoman", "%3."),
        };
        const int hangingBase = 240;
        int gap = spacingCfg.ListMarkerTextGap;
        int hangingVal = hangingBase + gap;
        int bodyIndent = spacingCfg.BodyFirstLine;

        for (int ilvl = 0; ilvl < 3; ilvl++)
        {
            int leftVal = bodyIndent + ilvl * 480;
            var (fmt, txt) = fmts[ilvl];
            var lvl = new Level { LevelIndex = ilvl };
            lvl.Append(new StartNumberingValue { Val = 1 });
            lvl.Append(new NumberingFormat { Val = ParseNumberFormat(fmt) });
            lvl.Append(new LevelText { Val = txt });
            lvl.Append(new LevelJustification { Val = LevelJustificationValues.Left });
            lvl.Append(new LevelSuffix { Val = LevelSuffixValues.Tab });
            lvl.Append(new PreviousParagraphProperties(
                new Tabs(new TabStop
                {
                    Val = TabStopValues.Number,
                    Position = leftVal,
                }),
                new Indentation
                {
                    Left = leftVal.ToString(),
                    Hanging = hangingVal.ToString(),
                }
            ));
            abs.Append(lvl);
        }
        return abs;
    }

    private static NumberFormatValues ParseNumberFormat(string fmt) => fmt switch
    {
        "decimal" => NumberFormatValues.Decimal,
        "chineseCountingThousand" => NumberFormatValues.ChineseCountingThousand,
        "upperLetter" => NumberFormatValues.UpperLetter,
        "lowerLetter" => NumberFormatValues.LowerLetter,
        "upperRoman" => NumberFormatValues.UpperRoman,
        "lowerRoman" => NumberFormatValues.LowerRoman,
        "bullet" => NumberFormatValues.Bullet,
        _ => NumberFormatValues.Decimal,
    };

    private static LevelSuffixValues ParseSuffix(string suff) => suff switch
    {
        "space" => LevelSuffixValues.Space,
        "nothing" => LevelSuffixValues.Nothing,
        _ => LevelSuffixValues.Tab,
    };

    private static void RemoveById<T>(Numbering root, Func<T, int?> idSelector, params int[] ids)
        where T : OpenXmlElement
    {
        var set = new HashSet<int>(ids);
        foreach (var el in root.Elements<T>().ToList())
        {
            var id = idSelector(el);
            if (id.HasValue && set.Contains(id.Value))
                el.Remove();
        }
    }

    // -----------------------------------------------------------------------
    // InjectSettings
    // -----------------------------------------------------------------------

    public static void InjectSettings(WordprocessingDocument doc)
    {
        var settingsPart = doc.MainDocumentPart!.DocumentSettingsPart;
        if (settingsPart?.Settings is null) return;
        foreach (var el in settingsPart.Settings.Elements<UpdateFieldsOnOpen>().ToList())
            el.Remove();
    }
}
