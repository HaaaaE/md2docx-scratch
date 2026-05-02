using System.Diagnostics;
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;

namespace Md2Docx;

/// <summary>
/// LaTeX → OMML via external pandoc binary.
/// Prebuild batches all formulas into one pandoc call; individual misses fall back to single calls.
/// </summary>
public class LatexConverter
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly string _pandocPath;
    private readonly Dictionary<(string latex, bool display), string> _cache = new();
    private readonly Action<string>? _warn;

    public LatexConverter(Action<string>? warn = null)
    {
        _warn = warn;
        _pandocPath = FindPandoc()
            ?? throw new InvalidOperationException(
                "未找到 pandoc 二进制。请先安装：\n" +
                "  Linux:  sudo apt install pandoc  /  sudo dnf install pandoc\n" +
                "  macOS:  brew install pandoc\n" +
                "  Windows: winget install pandoc  /  https://pandoc.org/installing.html");
    }

    private static string? FindPandoc()
    {
        var names = OperatingSystem.IsWindows()
            ? new[] { "pandoc.exe", "pandoc" }
            : new[] { "pandoc" };

        foreach (var name in names)
        {
            var pathDirs = Environment.GetEnvironmentVariable("PATH")?.Split(Path.PathSeparator) ?? [];
            foreach (var dir in pathDirs)
            {
                var full = Path.Combine(dir, name);
                if (File.Exists(full)) return full;
            }
        }
        return null;
    }

    /// <summary>Batch pre-convert all formulas with one pandoc invocation.</summary>
    public void Prebuild(List<(string latex, bool display)> formulas)
    {
        var unique = new List<(string latex, bool display)>();
        var seen = new HashSet<(string, bool)>();
        foreach (var (latex, display) in formulas)
        {
            var key = (latex.Trim(), display);
            if (key.Item1.Length == 0 || !seen.Add(key)) continue;
            unique.Add(key);
        }
        if (unique.Count == 0) return;

        var mdParts = unique.Select(f =>
            f.display ? $"$${f.latex}$$" : $"${f.latex}$");
        var mdText = string.Join("\n\n", mdParts) + "\n";

        var elements = PandocToOmml(mdText);
        if (elements is null) return;

        if (elements.Count != unique.Count)
            _warn?.Invoke($"pandoc batch returned {elements.Count} OMML elements, expected {unique.Count}");

        for (int i = 0; i < Math.Min(unique.Count, elements.Count); i++)
            _cache[(unique[i].latex.Trim(), unique[i].display)] = elements[i];
    }

    /// <summary>Get OMML XML string for a formula. Returns null on failure.</summary>
    public string? ToOmml(string latex, bool display)
    {
        latex = latex.Trim();
        if (latex.Length == 0) return null;

        var key = (latex, display);
        if (_cache.TryGetValue(key, out var cached))
            return cached;

        var wrapped = display ? $"$${latex}$$" : $"${latex}$";
        var elements = PandocToOmml(wrapped + "\n");
        if (elements is null || elements.Count == 0)
        {
            _warn?.Invoke($"pandoc returned no OMML for: {latex}");
            return null;
        }
        _cache[key] = elements[0];
        return elements[0];
    }

    /// <summary>Convert an OpenXmlElement from OMML XML string, inserting into target parent.</summary>
    /// <summary>Parse an OMML XML string into an OpenXmlElement that can be appended to a Paragraph.</summary>
    public static OpenXmlElement? ParseOmmlElement(string outerXml)
    {
        try
        {
            // Wrap in a temporary container so we can parse via InnerXml
            var wrapper = new OpenXmlUnknownElement("w", "body",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            { InnerXml = outerXml };
            return wrapper.FirstChild?.CloneNode(true);
        }
        catch
        {
            return null;
        }
    }

    private List<string>? PandocToOmml(string mdText)
    {
        var tmpPath = Path.Combine(Path.GetTempPath(), $"md2docx_{Guid.NewGuid():N}.docx");
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = _pandocPath,
                Arguments = $"-f markdown+tex_math_dollars+tex_math_double_backslash -t docx -o \"{tmpPath}\"",
                RedirectStandardInput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi);
            if (proc is null) return null;

            proc.StandardInput.Write(mdText);
            proc.StandardInput.Close();

            if (!proc.WaitForExit(60_000))
            {
                _warn?.Invoke("pandoc timed out");
                try { proc.Kill(); } catch { }
                return null;
            }

            if (proc.ExitCode != 0)
            {
                var stderr = proc.StandardError.ReadToEnd();
                _warn?.Invoke($"pandoc returned {proc.ExitCode}: {stderr.Trim()}");
                return null;
            }

            return ExtractOmmlFromDocx(tmpPath);
        }
        catch (Exception ex)
        {
            _warn?.Invoke($"pandoc invocation failed: {ex.Message}");
            return null;
        }
        finally
        {
            try { File.Delete(tmpPath); } catch { }
        }
    }

    private List<string>? ExtractOmmlFromDocx(string docxPath)
    {
        try
        {
            using var zip = ZipFile.OpenRead(docxPath);
            var entry = zip.GetEntry("word/document.xml");
            if (entry is null) return null;

            using var stream = entry.Open();
            var xdoc = XDocument.Load(stream);

            var elements = new List<string>();
            foreach (var p in xdoc.Descendants(W + "p"))
            {
                var oMathPara = p.Descendants(M + "oMathPara").FirstOrDefault();
                if (oMathPara is not null)
                {
                    ScrubNaryPlaceholders(oMathPara);
                    elements.Add(oMathPara.ToString());
                    continue;
                }
                var oMath = p.Descendants(M + "oMath").FirstOrDefault();
                if (oMath is not null)
                {
                    ScrubNaryPlaceholders(oMath);
                    elements.Add(oMath.ToString());
                }
            }
            return elements;
        }
        catch (Exception ex)
        {
            _warn?.Invoke($"failed to read pandoc output docx: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Clean up zero-width space placeholders pandoc puts in nary sub/sup elements.
    /// When pandoc generates \sum_{x} it writes supHide="on" but also puts U+200B in sup,
    /// causing Word to render a placeholder box.
    /// </summary>
    private static void ScrubNaryPlaceholders(XElement root)
    {
        var orderTags = new[] { "chr", "limLoc", "grow", "subHide", "supHide", "ctrlPr" }
            .Select(t => M + t).ToArray();

        foreach (var nary in root.Descendants(M + "nary").ToList())
        {
            var naryPr = nary.Element(M + "naryPr");

            foreach (var side in new[] { "sub", "sup" })
            {
                var elem = nary.Element(M + side);
                if (elem is null) continue;

                var text = string.Concat(elem.Descendants(M + "t").Select(t => t.Value));
                if (text.Replace("\u200b", "").Trim().Length > 0) continue;

                elem.RemoveNodes();

                naryPr ??= new XElement(M + "naryPr");
                if (naryPr.Parent is null)
                    nary.AddFirst(naryPr);

                var hideTag = M + $"{side}Hide";
                var hide = naryPr.Element(hideTag);
                if (hide is null)
                {
                    hide = new XElement(hideTag);
                    int targetIdx = Array.IndexOf(orderTags, hideTag);
                    var insertBefore = naryPr.Elements()
                        .FirstOrDefault(e => Array.IndexOf(orderTags, e.Name) > targetIdx);
                    if (insertBefore is not null)
                        insertBefore.AddBeforeSelf(hide);
                    else
                        naryPr.Add(hide);
                }
                hide.SetAttributeValue(M + "val", "1");
            }
        }
    }
}
