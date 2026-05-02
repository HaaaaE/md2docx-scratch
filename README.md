# Md2Docx

将 Markdown 转为 Word（`.docx`）的命令行工具，基于 [Markdig](https://github.com/xoofx/markdig) 与 [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)。

## 环境

- [.NET 10 SDK](https://dotnet.microsoft.com/download)

## 构建与运行

```bash
dotnet build
dotnet run -- project.md -o output.docx
```

（发布后可直接运行可执行文件，用法相同。）

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

未指定 `-c` 时使用内置默认样式。可参考仓库中的 `config.example.yaml` 自定义页面、字体、字号、颜色、段落间距与标题编号格式等。

## 依赖

- `DocumentFormat.OpenXml` — 生成文档
- `Markdig` — 解析 Markdown
- `YamlDotNet` — 读取配置文件

更细的实现说明见 `ARCHITECTURE.md`。
