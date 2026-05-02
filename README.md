# Md2Docx

将 Markdown 转为 Word（`.docx`）的命令行工具，基于 [Markdig](https://github.com/xoofx/markdig) 与 [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)。公式（`$…$` / `$$…$$`）通过外部 **pandoc** 转为 Word 中的 OMML。

## 环境

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- [pandoc](https://pandoc.org/installing.html) 已安装且在 `PATH` 中（程序启动时会检测；未安装会报错退出）

## 构建与运行

在项目根目录：

```bash
dotnet build
dotnet run -- project.md -o output.docx
```

若需要可单独分发的程序，可自行使用 `dotnet publish`。

## 用法

```text
md2docx <input.md> -o <output.docx> [-c config.yaml] [-v]
```

| 参数 | 说明 |
|------|------|
| `-o` / `--output` | 输出 `.docx` 路径（必填） |
| `-c` / `--config` | YAML 排版配置（可选） |
| `-v` / `--verbose` | 输出警告信息 |
| `-h` / `--help` | 显示帮助 |

未指定 `-c` 时使用内置默认样式。可参考仓库中的 `config.example.yaml` 自定义页面、字体、字号、颜色、段落间距与标题编号格式等。Markdown 中的本地图片等相对路径，相对于该 `.md` 文件所在目录解析。

## 依赖

- **pandoc**（外部程序，非 NuGet）：公式转换，须已在环境中安装（与「环境」一节一致）。
- `DocumentFormat.OpenXml` — 生成文档
- `Markdig` — 解析 Markdown
- `YamlDotNet` — 读取配置文件

更细的实现说明见 `ARCHITECTURE.md`。
