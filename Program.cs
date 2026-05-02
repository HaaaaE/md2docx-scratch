using Md2Docx;

string? inputPath = null;
string? outputPath = null;
string? configPath = null;
bool verbose = false;

for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "-o" or "--output" when i + 1 < args.Length:
            outputPath = args[++i];
            break;
        case "-c" or "--config" when i + 1 < args.Length:
            configPath = args[++i];
            break;
        case "-v" or "--verbose":
            verbose = true;
            break;
        case "-h" or "--help":
            PrintUsage();
            return 0;
        default:
            if (!args[i].StartsWith('-') && inputPath is null)
                inputPath = args[i];
            else
            {
                Console.Error.WriteLine($"未知参数: {args[i]}");
                PrintUsage();
                return 1;
            }
            break;
    }
}

if (inputPath is null || outputPath is null)
{
    Console.Error.WriteLine("缺少必要参数。");
    PrintUsage();
    return 1;
}

if (!File.Exists(inputPath))
{
    Console.Error.WriteLine($"输入文件不存在: {inputPath}");
    return 2;
}

if (configPath is not null && !File.Exists(configPath))
{
    Console.Error.WriteLine($"配置文件不存在: {configPath}");
    return 2;
}

Action<string>? warn = verbose
    ? msg => Console.Error.WriteLine($"WARNING: {msg}")
    : null;

var cfg = Config.Load(configPath);
var mdText = File.ReadAllText(inputPath);

MdToDocxConverter converter;
try
{
    converter = new MdToDocxConverter(inputPath, cfg, warn);
}
catch (InvalidOperationException ex)
{
    Console.Error.WriteLine(ex.Message);
    return 3;
}

converter.Convert(mdText);
converter.Save(outputPath);
Console.WriteLine($"已生成: {outputPath}");
return 0;

static void PrintUsage()
{
    Console.Error.WriteLine("用法: md2docx <input.md> -o <output.docx> [-c config.yaml] [-v]");
}
