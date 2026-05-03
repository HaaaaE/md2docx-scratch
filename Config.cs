using System.Collections;
using System.Globalization;
using System.Reflection;
using YamlDotNet.Serialization;

namespace Md2Docx;

public record PageConfig
{
    public int Width { get; init; } = 12240;       // twips, 21.59cm Letter
    public int Height { get; init; } = 15840;      // twips, 27.94cm
    public int Margin { get; init; } = 1134;       // twips, 2.00cm
    public int HeaderDist { get; init; } = 720;    // twips, 1.27cm
    public int FooterDist { get; init; } = 720;
    public int Gutter { get; init; } = 567;        // twips, 1.00cm
}

public record FontsConfig
{
    public string Latin { get; init; } = "Times New Roman";
    public string CjkBody { get; init; } = "楷体";
    public string CjkH1 { get; init; } = "华文楷体";
    public string Code { get; init; } = "Consolas";
}

public record SizesConfig
{
    public int Body { get; init; } = 24;   // half-points, 12pt
    public int H1 { get; init; } = 40;    // half-points, 20pt
    public int Code { get; init; } = 20;  // half-points, 10pt
}

public record ColorsConfig
{
    public string H1 { get; init; } = "3C7BC8";
    public string H4 { get; init; } = "0F4761";
}

public record SpacingConfig
{
    public int BodyBefore { get; init; } = 0;
    public int BodyAfter { get; init; } = 120;
    public int BodyLine { get; init; } = 240;
    public string BodyLineRule { get; init; } = "auto";
    public int BodyFirstLine { get; init; } = 480;
    public int BodyFirstLineChars { get; init; } = 200;

    public int H1Before { get; init; } = 0;
    public int H1After { get; init; } = 120;
    public int H2Before { get; init; } = 0;
    public int H2After { get; init; } = 120;
    public int H3Before { get; init; } = 0;
    public int H3After { get; init; } = 120;
    public int H4Before { get; init; } = 80;
    public int H4After { get; init; } = 40;

    public int ListBefore { get; init; } = 0;
    public int ListAfter { get; init; } = 60;
    /// <summary>Extra twips added to list <c>w:hanging</c> only: widens marker tab field so bullets/numbers sit farther from text; list text start still uses <see cref="BodyFirstLine"/> + ilvl×480.</summary>
    public int ListMarkerTextGap { get; init; } = 80;
    public int ListLeft { get; init; } = 480;
    public int ListHanging { get; init; } = 480;

    public int BlockquoteBefore { get; init; } = 0;
    public int BlockquoteAfter { get; init; } = 120;
    public int BlockquoteLeft { get; init; } = 567;

    public int CodeBlockBefore { get; init; } = 0;
    public int CodeBlockAfter { get; init; } = 0;
    public int CodeBlockLeft { get; init; } = 284;
    public int CodeBlockRight { get; init; } = 284;
}

public record NumberingConfig
{
    public string H1NumFmt { get; init; } = "chineseCountingThousand";
    public string H1LvlText { get; init; } = "第%1章";
    public string H1Suff { get; init; } = "space";
    public string SubSep { get; init; } = ".";
    public string SubTrailing { get; init; } = "";
    public string SubSuff { get; init; } = "space";
}

public record CaptionConfig
{
    public string FigLabel { get; init; } = "图";
    public string TabLabel { get; init; } = "表";
    public string Sep { get; init; } = "-";
    public bool IncludeChapter { get; init; } = true;
}

public record BulletConfig
{
    public string L0 { get; init; } = "\u00B7";
    public string L1 { get; init; } = "\u00B7";
    public string L2 { get; init; } = "\u00B7";
}

public record OrderedListConfig
{
    public string L0Fmt { get; init; } = "decimal";
    public string L0Suffix { get; init; } = ".";
    public string L1Fmt { get; init; } = "lowerLetter";
    public string L1Suffix { get; init; } = ")";
    public string L2Fmt { get; init; } = "lowerRoman";
    public string L2Suffix { get; init; } = ".";
}

public record Config
{
    public PageConfig Page { get; init; } = new();
    public FontsConfig Fonts { get; init; } = new();
    public SizesConfig Sizes { get; init; } = new();
    public ColorsConfig Colors { get; init; } = new();
    public SpacingConfig Spacing { get; init; } = new();
    public NumberingConfig Numbering { get; init; } = new();
    public CaptionConfig Caption { get; init; } = new();
    public BulletConfig Bullet { get; init; } = new();
    public OrderedListConfig OrderedList { get; init; } = new();

    private static readonly IDeserializer YamlRootDeserializer =
        new DeserializerBuilder().IgnoreUnmatchedProperties().Build();

    public static Config Load(string? path)
    {
        if (path is null) return new Config();

        var yaml = File.ReadAllText(path);
        if (string.IsNullOrWhiteSpace(yaml))
            return new Config();

        var root = YamlRootDeserializer.Deserialize<Dictionary<object, object?>>(yaml)
            ?? new Dictionary<object, object?>();

        return MergeFromYamlRoot(root);
    }

    private static Config MergeFromYamlRoot(Dictionary<object, object?> root)
    {
        object merged = new Config();
        foreach (var prop in typeof(Config).GetProperties(BindingFlags.Instance | BindingFlags.Public))
        {
            if (!TryGetMappingValue(root, prop.Name, out var node))
                continue;
            object currentNested = prop.GetValue(merged)!;
            object updated = MergeNestedRecord(currentNested, node);
            if (!ReferenceEquals(currentNested, updated))
                merged = SetPropertyOnRecordClone(merged, prop, updated);
        }
        return (Config)merged;
    }

    private static bool TryGetMappingValue(Dictionary<object, object?> dict, string key, out object? value)
    {
        foreach (var kv in dict)
            if (string.Equals(kv.Key.ToString()?.Trim(), key, StringComparison.OrdinalIgnoreCase))
            {
                value = kv.Value;
                return true;
            }
        value = null;
        return false;
    }

    private static object MergeNestedRecord(object record, object? yamlSection)
    {
        if (!TryEnumerateStringKeyedScalars(yamlSection, out var entries))
            return record;

        var type = record.GetType();
        object result = record;
        foreach (var (k, yamlValue) in entries)
        {
            var nestedProp = GetPropertyIgnoringCase(type, k);
            if (nestedProp is null) continue;

            if (!TryCoerceForProperty(nestedProp.PropertyType, yamlValue, out var coerced))
                continue;
            result = SetPropertyOnRecordClone(result, nestedProp, coerced);
        }
        return result;
    }

    private static PropertyInfo? GetPropertyIgnoringCase(Type type, string name)
    {
        foreach (var p in type.GetProperties(BindingFlags.Instance | BindingFlags.Public))
            if (string.Equals(p.Name, name.Trim(), StringComparison.OrdinalIgnoreCase))
                return p;
        return null;
    }

    /// <summary>YamlDotNet 将嵌套 mapping 转成字典；逐项合并到 record。</summary>
    private static bool TryEnumerateStringKeyedScalars(object? node, out IEnumerable<(string Key, object? Value)> pairs)
    {
        if (node is not IDictionary raw)
        {
            pairs = Array.Empty<(string, object?)>();
            return false;
        }
        pairs = EnumeratePairs(raw);
        return true;
    }

    private static IEnumerable<(string Key, object? Value)> EnumeratePairs(IDictionary raw)
    {
        foreach (DictionaryEntry kv in raw)
            yield return (kv.Key.ToString()!, kv.Value);
    }

    private static bool TryCoerceForProperty(Type targetType, object? yamlValue, out object? coerced)
    {
        if (yamlValue is null || targetType == typeof(void))
        {
            coerced = null;
            return false;
        }

        if (targetType == typeof(int))
        {
            try
            {
                coerced = ConvertToInt32(yamlValue);
                return true;
            }
            catch (InvalidCastException)
            {
                coerced = null;
                return false;
            }
        }

        if (targetType == typeof(string))
        {
            if (yamlValue is string s)
            {
                coerced = s;
                return true;
            }
            coerced = yamlValue.ToString();
            return coerced is not null;
        }

        if (targetType == typeof(bool))
        {
            if (yamlValue is bool b) { coerced = b; return true; }
            if (yamlValue is string bs && bool.TryParse(bs, out var bv)) { coerced = bv; return true; }
            coerced = null;
            return false;
        }

        coerced = null;
        return false;
    }

    private static int ConvertToInt32(object yamlValue)
    {
        switch (yamlValue)
        {
            case int i: return i;
            case short sht: return sht;
            case long lng: return checked((int)lng);
            case uint ui: return checked((int)ui);
            case ulong ulg: return checked((int)ulg);
            case byte bt: return bt;
            case double dbl: return (int)Math.Round(dbl);
            case float fl: return (int)Math.Round(fl);
            case decimal dc: return (int)Math.Round(dc);
            case string ss when int.TryParse(ss, NumberStyles.Integer, CultureInfo.InvariantCulture, out var si):
                return si;
            default:
                return Convert.ToInt32(yamlValue, CultureInfo.InvariantCulture);
        }
    }

    private static object SetPropertyOnRecordClone(object record, PropertyInfo prop, object? value)
    {
        var cloneMethod =
            record.GetType().GetMethod("<Clone>$", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException($"{record.GetType().FullName} 不是可被合并的 record 类型");
        object clone = cloneMethod.Invoke(record, null)!;
        prop.SetValue(clone, value);
        return clone;
    }
}
