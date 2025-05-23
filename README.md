# markdown2article

本工具旨在将特定格式的 Markdown 文档转换为符合“天津理工大学本科毕业设计说明书（毕业论文）”规范的 Word (.docx) 文档。这是一个初版工具，部分功能仍在完善中。

## 功能特性

*   **结构转换：** 能够识别 Markdown 中的特定标题结构（摘要、ABSTRACT、正文章节、参考文献、致谢）并转换为论文要求的格式。
*   **样式应用：** 尽力根据论文规范设置字体（宋体、Times New Roman）、字号、行距、对齐方式等。
*   **图片处理：**
    *   转换标准 Markdown 图片，并提取图片标题。
    *   通过 `mmdc` 工具将 Mermaid 图表转换为图片，并提取图表标题。
*   **表格处理：** 转换 Markdown 表格，并提取表格标题。
*   **参考文献：** 将 Markdown 脚注形式的参考文献转换为论文要求的格式，并处理文内引用（目标是上角标）。
*   **页眉页脚：** （目标功能）根据规范设置页眉和页码。

## 当前状态与已知问题

*   **[重要] 数学公式：** 目前**不支持** Markdown 内嵌数学公式（如 LaTeX 语法）的转换。
*   **[重要] 目录生成：** 目录生成功能目前**无法正常使用**。
*   **样式细节：** 虽然尽力符合规范，但部分 Word 复杂样式（如固定行距20磅与1.25倍行距的精确切换、页眉页脚的精确实现）可能存在偏差，需要用户在生成后进行微调。
*   **错误处理：** 错误处理和提示信息可能不够完善。

## 依赖项

### Python 库
*   `beautifulsoup4` (bs4)
*   `Markdown` (Python Markdown parser)
*   `python-docx`

建议创建一个 `requirements.txt` 文件并使用 pip 安装：
```
beautifulsoup4
Markdown
python-docx
```
(其他库如 `argparse`, `os` 等通常是 Python 标准库，无需额外安装)

### 命令行工具
*   **`mmdc`** (Mermaid CLI): 用于将 Mermaid 图表转换为图片。
    需要通过 npm (Node.js 包管理器) 安装。如果你尚未安装 Node.js 和 npm, 请先安装它们。
    安装 `mmdc`：
    ```bash
    npm install -g @mermaid-js/mermaid-cli
## 安装与准备

1.  确保你已安装 Python 3.x。
2.  安装 Node.js 和 npm。
3.  安装 `mmdc`：
    ```bash
    npm install -g @mermaid-js/mermaid-cli
    ```
4.  安装所需的 Python 库：
    ```bash
    pip install beautifulsoup4 Markdown python-docx
    ```
5.  将 `markdown2artical.py` 脚本放置在你的工作目录或一个可通过 PATH 访问的目录。

## 使用方法

```bash
python markdown2artical.py <input_markdown_file.md> <output_word_file.docx>
```

**示例：**

```bash
python markdown2artical.py my_paper.md my_paper_final.docx
```

脚本将读取 `my_paper.md` 文件，按照预设的规则和论文格式要求进行转换，并生成 `my_paper_final.docx` 文件。

## 输入 Markdown 格式约定

为了工具能够正确解析和转换，你的 Markdown 文档需要遵循以下约定：

### 1. 整体结构与一级标题

文档必须按照以下顺序组织，并使用一级标题 (`#`)：

1.  `# 摘要` (中文摘要)
2.  `# ABSTRACT` (英文摘要)
3.  `# 第X章 章节标题` (例如: `# 第一章 引言`, `# 第二章 相关工作`) - 正文部分
4.  `# 参考文献`
5.  `# 致谢`

### 2. 正文标题层级

*   **一级标题 (章标题):** 格式为 `第X章 章节名` (例如: `# 第一章 引言`)。章序号与标题之间应空一个汉字。
*   **二级标题 (节标题):** 格式为 `## X.Y 节标题` (例如: `## 1.1 研究背景`)。节序号与标题之间应空一个西文字符。
*   **三级标题 (小节标题):** 格式为 `### X.Y.Z 小节标题` (例如: `### 1.1.1 国内研究现状`)。小节序号与标题之间应空一个西文字符。
*   **四级标题:** 格式为 `#### 1. 标题内容` (阿拉伯数字后跟点和空格)。不计入目录。
*   **五级标题:** 格式为 `##### (1) 标题内容` (括号数字后跟空格)。不计入目录。

### 3. 图片

*   **普通图片：**
    直接使用 Markdown 的图片语法，图题包含在 `alt` 文本中，并以 "图X.Y " 开头。
    ```markdown
    ![图2.1 某结构示意图](path/to/your/image.png)
    ```
    工具会提取 "图2.1 某结构示意图" 作为图题，并按规范放置。

*   **Mermaid 图表：**
    使用 `mermaid` 代码块。图题必须写在 `mermaid` 代码块的第一行，并以 `%%` (Mermaid 注释) 开头，后接 "图X.Y "。
    
    ```
    %%图3.1 某功能流程图
    graph TD
        A[开始] --> B{条件?};
        B -- 是 --> C[操作1];
        B -- 否 --> D[操作2];
        C --> E[结束];
        D --> E;
    ```
    工具会提取 "图3.1 某功能流程图" 作为图题，并将 Mermaid 代码转换为图片。

### 4. 表格

*   使用标准 Markdown 表格语法。
*   表格标题必须放在第一行表头的第一个单元格内，并用中括号 `[]` 包裹，以 "表X.Y " 开头。
    ```markdown
    | [表2.1 典型虚拟化环境的物理前缀特征]虚拟化平台 | 物理地址前缀                           |
    | ---------------------------------------------- | -------------------------------------- |
    | Parallels                                      | 00:1C:42                               |
    | VirtualBox                                     | 08:00:27                               |
    | VMware                                         | 00:05:69, 00:0C:29, 00:1C:14, 00:50:56 |
    | Xen                                            | 00:16:E3                               |
    ```
    工具会提取 "[表2.1 典型虚拟化环境的物理前缀特征]" 作为表题，并按规范放置。

### 5. 参考文献

*   **定义：** 使用 Markdown 的脚注（footnote）格式定义参考文献。脚注标签即为引用序号。
    ```markdown
    [^1]: 作者. 文献题名[文献类型标识]. 出版地: 出版者, 出版年: 起止页码. 获取和访问路径.
    [^2]: Collembolan assemblages in forest soils: the effects of forest type and soil P [J]. Pedobiologia, 2003, 47(1): 1-10. https://example.com/ref2. 2024-01-15
    ```
    注意：脚注内容本身的格式应已基本符合论文要求。工具主要负责转换序号格式。

*   **引用：** 在正文中，使用 Markdown 脚注引用方式。
    ```markdown
    ...正如文献[文献[^1]和文献[^2]所述...
    ```
    工具会尝试将这些引用转换为 Word 中的上角标格式 `[1]`、`[2]`。

## 输出 Word 文档格式目标 (基于草案)

本工具旨在生成符合以下主要规范的 Word 文档：

*   **纸张与页边距：** A4，上/下/左/右 2.5cm。
*   **默认字体：**
    *   中文：小四号宋体
    *   英文/数字/符号：Times New Roman
*   **行间距：** 默认 1.25 倍行距。
*   **页眉：** 从正文开始至致谢结束，“天津理工大学2025届本科毕业设计说明书（毕业论文）”，宋体五号居中，带页眉线。
*   **页码：** 从正文开始，阿拉伯数字，宋体小五号，居中。

### 各部分格式细节：

*   **中文摘要：**
    *   标题“摘要”：黑体三号居中（“摘”“要”间空一汉字），后空一行。
    *   正文：楷体四号，段首缩进，行距固定值20磅。
*   **英文摘要 (ABSTRACT)：**
    *   中文摘要后空两行，打印英文标题，空一行英文副标题。
    *   空一行，标题“ABSTRACT”：三号加粗居中。
    *   正文：四号，段首缩进，行距固定值20磅。
*   **正文：**
    *   **章标题** (如“第一章 引言”)：三号黑体居中，每章另起一页。标题1.5倍行距。章序号与标题间空一汉字。
    *   **节标题** (如“1.1 XXX”)：小三号黑体居中。节序号与标题间空一西文字符。
    *   **小节标题** (如“1.1.1 XXX”)：四号黑体左对齐。小节序号与标题间空一西文字符。
    *   **图题：** 中文五号楷体，图序号 Times New Roman，图下方居中。图号按章编号 (如图3.2)。
    *   **表题：** 五号楷体，表序号 Times New Roman，表上方居中。表号按章编号 (如表5.4)。
*   **参考文献：**
    *   另起一页，标题“参考文献”：三号黑体居中。
    *   后空一行，正文：小四号宋体，左对齐，悬挂缩进，1.25 倍行距。序号格式为 `[1]`。
*   **致谢：**
    *   标题“致谢”：三号黑体居中（“致”“谢”间空两汉字）。
    *   正文：小四号宋体 (中文)，Times New Roman (英文/数字)，1.25 倍行距。

## 未来工作 / TODO

*   [ ] **实现数学公式转换** (例如，使用 MathJax 或 Pandoc 的能力)。
*   [ ] **修复并完善目录生成功能** (深入到三级标题)。
*   [ ] 增强对 Word 样式的控制，减少手动调整。
*   [ ] 完善页眉页脚的自动生成与切换。
*   [ ] 增加更详细的错误报告和用户反馈。
*   [ ] 允许用户通过配置文件自定义部分样式参数。

## 注意事项

*   本工具为初版，转换结果可能需要用户在 Word 中进行二次审阅和调整，以确保完全符合最终的论文提交要求。
*   请严格按照“输入 Markdown 格式约定”编写 Markdown 文档，否则可能导致转换失败或格式错乱。
*   由于 Markdown 本身的限制以及 `python-docx` 库对 Word 文档的操作能力，某些非常精细的排版效果可能难以完美实现自动化。

希望这个工具能对你有所帮助！
