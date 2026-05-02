# Md2Docx 技术架构

Markdown → DOCX 转换器。通过 Word 样式注入（而非 direct formatting）实现格式控制，输出的文档可在 Word 样式面板中直接修改样式并全局生效。

## 运行时依赖

| 依赖 | 用途 |
|---|---|
| `DocumentFormat.OpenXml` 3.5.x | OOXML 强类型 API，构造 docx 内部 XML |
| `Markdig` 1.1.x | Markdown 解析，生成 AST |
| `YamlDotNet` 16.x | YAML 配置文件解析 |
| **pandoc**（外部二进制） | LaTeX 公式 → OMML 转换，必须在 PATH 中可用 |

## 文件职责

```
Md2Docx.csproj          项目定义
Program.cs              CLI 入口（顶级语句；-c/--config 指向 YAML）
Config.cs               配置 record 定义 + YAML 加载
config.example.yaml     与 Config 属性树一致的示例文件（可复制改）
StyleDefs.cs            样式规范字典（数据层）
OoxmlHelpers.cs         OOXML 元素构造工具函数
PageSetup.cs            页面尺寸 / 页边距 / 节属性
StyleInjector.cs        样式 / 编号 / 设置注入到 docx
LatexConverter.cs       pandoc 调用 + OMML 缓存
InlineRenderer.cs       Markdig inline AST → Word Run 序列
MdToDocxConverter.cs    主转换器：解析 → 预扫描 → AST 遍历 → 渲染
```

### 依赖方向（单向，无循环）

```
Program.cs
  └→ MdToDocxConverter
       ├→ InlineRenderer
       │    ├→ OoxmlHelpers
       │    └→ LatexConverter
       ├→ StyleInjector
       │    ├→ OoxmlHelpers
       │    └→ StyleDefs
       ├→ PageSetup
       └→ Config
```

## 数据流

```
CLI args
  │
  ├─ Config.Load(yaml)          解析 YAML，返回不可变 Config record
  │
  └─ new MdToDocxConverter(path, cfg)
       │
       │  构造函数内部：
       │    WordprocessingDocument.Create(MemoryStream)
       │    → ClearDefaultBody → SetupPage
       │    → InjectNumbering → InjectStyles → InjectSettings
       │    → new LatexConverter
       │    → 配置 Markdig Pipeline
       │
       └─ converter.Convert(mdText)
            │
            ├─ Markdig.Parse → MarkdownDocument (AST)
            ├─ CollectFormulas → LatexConverter.Prebuild (批量 pandoc)
            ├─ PrecomputeCaptionNumbers (图表编号预扫描)
            ├─ RenderBlocks (遍历 AST 子块，逐块渲染)
            └─ FlushPending (表注 look-back 缓冲刷出)
```

## 核心设计决策

### 样式注入 vs 直接格式化

每个段落只设 `w:pStyle`，不写 direct `w:pPr`/`w:rPr`。所有排版参数（字体、字号、间距、缩进）在 `StyleInjector.InjectStyles` 中写入 `styles.xml`。好处：

- Word 中修改样式即可全局生效
- 输出 XML 更简洁
- 格式参数集中在 `StyleDefs.Build` 一处管理

### 样式规范的中间表示

`StyleSpec` record 用 `Dictionary<string, object?>` 存储 pPr/rPr 参数（key 如 `"before"`、`"bold"`、`"ascii"`）。`OoxmlHelpers.BuildPPr`/`BuildRPr` 负责把这个字典翻译为 Open XML SDK 强类型对象。

这层间接是为了让 `StyleDefs.Build` 保持纯数据声明，不直接依赖 Open XML SDK 的类型层次。新增样式只需在字典里加 entry。

### Markdig AST 遍历

Markdig 产生树形 AST（`MarkdownDocument` → `Block` → `Inline`），遍历方式：

- **Block 层**：`foreach (var block in document)` + `switch` 按类型分发
- **Inline 层**：递归 `WalkInline`，通过参数传递当前 bold/italic/strike 状态

#### display math 的特殊处理

Markdig 对单行 `$$...$$` 不产生 `MathBlock`，而是将其解析为 `ParagraphBlock` 内的 `MathInline`（`DelimiterCount == 2`）。`FlushPending` 中通过 `InlineRenderer.GetDisplayMath` 检测这种情况，将其路由到 `RenderMathBlock`。只有 `$$` 独占多行（类似 fenced code block 语法）才会产生真正的 `MathBlock`。

display math 的判定集中在 `InlineRenderer.IsDisplayMath(MathInline)`。预扫描公式缓存和实际渲染必须共用这个方法，避免 `LatexConverter` 的 `(latex, display)` cache key 在预构建和渲染阶段不一致。

### 表注 look-back 缓冲

遍历顶层块时，`ParagraphBlock` 不立即渲染，而是存入 `_pending`。遇到 `Table` 时将 `_pending` 作为表注标题渲染；遇到其他块时将 `_pending` 作为普通段落刷出。这实现了「表格上方的段落自动成为表注」的语义。

### 图表自动编号

`PrecomputeCaptionNumbers` 在渲染前扫描 AST，记录每张图/表所在的章节号和章内序号。渲染时通过 `SEQ` / `STYLEREF` Word 域插入编号，支持 Word 自动更新。

### LaTeX 公式管线

`LatexConverter` 调用外部 pandoc 将 LaTeX 转为 OMML：

1. **批量预转换**：`Prebuild` 将所有公式拼成一个 markdown 文件，一次 pandoc 调用生成 docx，解压提取 `m:oMath`/`m:oMathPara` 元素缓存
2. **单条回退**：cache miss 时逐条调 pandoc
3. **OMML 桥接**：pandoc 输出的 docx 用 `ZipFile` + `XDocument` 解析，提取的 XML 字符串通过 `OpenXmlUnknownElement.InnerXml` 桥接为 SDK 对象
4. **nary 清理**：`ScrubNaryPlaceholders` 移除 pandoc 在 `\sum`/`\prod` 等 nary 算子的 sub/sup 中插入的零宽空格占位符

### 行内代码格式独立性

`CodeInline` 渲染时使用固定的 rPr（等宽字体 + 灰色底纹），不继承外层的 bold/italic/strike 状态。这确保引用块等斜体环境中的行内代码保持正体。

## 配置系统

`Config` 由 6 个 `record` 组成，每个字段有默认值。`Config.Load` 将 YAML 根节点反序列化为嵌套字典，再按与 `Config` 相同的属性名（PascalCase，键名比对不区分大小写）逐项合并进默认值树；对每个嵌套 section 也用反射调用 record 的 `<Clone>$` 做 immutable 覆盖。未写入的顶层或小节沿用代码中的默认 `record`。

YAML 结构与 `Config`/`PageConfig`/… 的公有属性完全一致，等价于把整个配置对象写出来；可以只提供需要改写的顶层小节或小节内字段。

配置小节对应关系：

| YAML 根键（示例） | record | 控制内容 |
|---|---|---|
| `Page` | `PageConfig` | 纸张尺寸、页边距、装订线 |
| `Fonts` | `FontsConfig` | 西文/中文/代码字体 |
| `Sizes` | `SizesConfig` | 正文/标题/代码字号 |
| `Colors` | `ColorsConfig` | 标题颜色 |
| `Spacing` | `SpacingConfig` | 各元素段前段后、缩进 |
| `Numbering` | `NumberingConfig` | 一级标题编号格式 |

## 编号系统

`StyleInjector.InjectNumbering` 注入 3 个 `AbstractNum`：

| ID | 类型 | 用途 |
|---|---|---|
| 100 | multilevel (4 级) | 标题编号：h1 中文计数 / h2-h4 十进制多级 |
| 101 | hybridMultilevel (3 级) | 无序列表：中点符号 |
| 102 | hybridMultilevel (3 级) | 有序列表：数字 / 小写字母 / 小写罗马 |

ID 从 100 起步以避免与 Word 内置编号冲突。注入是幂等的（先删旧的再加）。

## 扩展指南

### 新增样式

1. 在 `StyleDefs.Build` 中添加新的字典 entry
2. 渲染代码中用 `Oxml.MakePara(_body, "new_key", _styleDefs)` 引用

### 新增 block 类型

1. 在 `MdToDocxConverter.RenderBlocks` 的 `switch` 中添加 `case`
2. 注意：Markdig 的类型继承关系影响 case 顺序——`MathBlock` 继承自 `FencedCodeBlock` 继承自 `CodeBlock`，子类必须在父类之前

### 新增 inline 类型

1. 在 `InlineRenderer.WalkInline` 的 `switch` 中添加 `case`
2. 如果需要特殊的 rPr，在同文件中添加对应的 builder

### 修改页面布局

修改 `PageConfig` 默认值，或通过 YAML 里的 `Page` 小节覆盖。`PageSetup.Setup` 会自动应用。

### 修改编号格式

修改 `NumberingConfig` 默认值或 YAML 里的 `Numbering` 小节。`StyleInjector.BuildHeadingAbstractNum` 使用这些值构造 multilevel 编号。
