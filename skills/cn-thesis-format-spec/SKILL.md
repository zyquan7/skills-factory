---
name: cn-thesis-format-spec
description: Enforce Chinese university thesis and translated academic paper formatting rules (Rules 2.1 to 2.16), including SimHei/SimSun/KaiTi/Times New Roman mixed-script typography, native Word OMML (<m:oMath>) inline and display equations, Chinese point sizes (XiaoEr to WuHao), full-width three-line tables (SanXianBiao), GB/T 7714-2015 references, and Word Outline Levels (<w:outlineLvl>) for headings, references, and appendices. Use whenever formatting a Chinese academic report, thesis, or translated paper in Word (.docx).
---
# CN Thesis Format Spec

A deterministic typography, native equation, outline hierarchy, and table layout specification implementing Chinese academic document standards (Rules 2.1 through 2.16) for Word (`.docx`) generation.

## When to Use

- Formatting a Chinese translated paper, literature review, or graduation thesis into `.docx`.
- Enforcing native Word OMML (`<m:oMath>`) equations, Word paragraph outline levels (`<w:outlineLvl>`), and standard Chinese three-line tables (`简明三线表`) alongside mixed Chinese/English font rules (`黑体`, `宋体`, `楷体` + `Times New Roman`).

## Steps

### Step 1: Map Chinese Font Size Names to Exact Point Values

Always use the exact point (`Pt`) values below:

| Chinese Size Name | Exact Point Size (`Pt`) | Half-Point Value (`w:sz`) | Primary Usage in Spec |
| :--- | :---: | :---: | :--- |
| **小二号 (XiaoEr)** | `18.0 pt` | `36` | Rule 2.1: Cover Chinese Title (黑体小二号加粗) |
| **三号 (SanHao)** | `16.0 pt` | `32` | Rule 2.1 Cover Foreign Title, Rule 2.2 CN Title, Rule 2.4 EN Title, Rule 2.6 TOC Title |
| **小三号 (XiaoSan)** | `15.0 pt` | `30` | Rule 2.1: Cover metadata fields (姓名/学号/专业/学院/指导教师 宋体小三号) |
| **四号 (SiHao)** | `14.0 pt` | `28` | Rule 2.2/2.4 Author Name, Rule 2.6 Level-1 TOC, Rule 2.7 H1 Heading, Rule 2.8/2.9 Refs & Appendix Title |
| **小四号 (XiaoSi)** | `12.0 pt` | `24` | Rule 2.2-2.5 Abstracts & Keywords, Rule 2.6 L2/L3 TOC, Rule 2.7 H2/H3 & Body Text, Rule 2.8 Refs, Rule 2.9 Appendix Body |
| **五号 (WuHao)** | `10.5 pt` | `21` | Rule 2.2/2.4 Affiliations, Rule 2.6 Blank Line, Rule 2.10 Figure Captions, Rule 2.11 Table Captions & Cells |

### Step 2: Apply Mixed Chinese/English Font Binding and Native OMML Equations

1. **Dual-Script Text Runs**: Set `w:ascii="Times New Roman"` and `w:hAnsi="Times New Roman"` on all runs, and set `w:eastAsia` to `黑体` (`SimHei`), `宋体` (`SimSun`), or `楷体` (`KaiTi`).
2. **Native Word OMML Math (`<m:oMath>`) for Both Inline and Display Formulas**:
   - Never render mathematical variables with subscripts/superscripts (e.g. $x_p$, $I_{\mathrm{test}}$, $\mathcal{I}_{\mathrm{train}}$, $\mathbb{R}^{D 	imes r}$, $\lambda_i$) as plain text with underscores (`_`).
   - Every inline `$...$` expression and every numbered display equation must be constructed as a native OpenXML `<m:oMath>` object (`xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`) using `<m:sSub>`, `<m:sSup>`, `<m:sSubSup>`, `<m:f>`, `<m:nary>`, and `<m:rad>` with `Cambria Math` font binding.

### Step 3: Enforce Authentic Three-Line Table (`简明三线表`, Rule 2.11)

To guarantee that three-line tables render identically in Microsoft Word, WPS Office, and LibreOffice without missing borders or cramped column wrapping:

- **Full-Width Table Sizing**: Strip default `<w:tblStyle>` from `<w:tblPr>`, set `<w:tblW w:w="5000" w:type="pct"/>` (100% of the `16.0 cm` printable width), and assign proportional column widths (`cell.width = Cm(...)`) based on each column's text length so metric headers (`I-AUROC`, `P-AUROC`, `PRO`) never wrap mid-word.
- **Exact Three-Line Border Weights**:
  - **Top rule (顶线)**: `1.5 pt` (`w:val="single" w:sz="12" w:color="000000"`) on `<w:tblBorders>` top and every top header cell's `<w:tcBorders>`.
  - **Spanner sub-rule (复合表头分组线, optional)**: `0.5 pt` (`w:val="single" w:sz="4" w:color="000000"`) under merged dataset group headers (such as `MVTec-AD` and `VisA`).
  - **Header bottom rule (栏目线)**: `0.75 pt` (`w:val="single" w:sz="6" w:color="000000"`) on every bottom header cell's `<w:tcBorders>`.
  - **Bottom rule (底线)**: `1.5 pt` (`w:val="single" w:sz="12" w:color="000000"`) on `<w:tblBorders>` bottom and every last data row cell's `<w:tcBorders>`.
  - **Zero vertical or interior horizontal lines**: Explicitly set `w:left`, `w:right`, `w:insideH`, and `w:insideV` to `w:val="nil"` on both `<w:tblBorders>` and every individual cell's `<w:tcBorders>`.

### Step 4: Enforce Native Word Outline Levels (`大纲级别`, Rule 2.16)

Every structural heading in the document MUST carry a native Word paragraph outline level (`<w:pPr><w:outlineLvl w:val="..."/></w:pPr>`) so that Microsoft Word and WPS Office display a complete, collapsible hierarchy in the Navigation Pane (`导航窗格`) and support automatic TOC field generation.

> **OpenXML Indexing Note**: In WordprocessingML (`<w:outlineLvl w:val="N"/>`), `w:val` is **0-based** (`0` = 1 级大纲级别, `1` = 2 级大纲级别, `2` = 3 级大纲级别, ..., `8` = 9 级大纲级别; omitting `<w:outlineLvl>` = 正文文本 / Body Text).

| Structural Element | Condition | Word UI Outline Level (`大纲级别`) | OpenXML `<w:outlineLvl>` |
| :--- | :---: | :---: | :---: |
| **正文一级标题** (e.g., `1 绪论`) | 必设 | **1 级** (`outline_level = 1`) | `<w:outlineLvl w:val="0"/>` |
| **正文二级标题** (e.g., `1.1 研究背景`) | 必设 | **2 级** (`outline_level = 2`) | `<w:outlineLvl w:val="1"/>` |
| **正文三级标题** (e.g., `1.1.1 问题定义`) | 必设 | **3 级** (`outline_level = 3`) | `<w:outlineLvl w:val="2"/>` |
| **正文第 $k$ 级标题** ($4 \le k \le 9$) | 按需依次类推 | **$k$ 级** (`outline_level = k`) | `<w:outlineLvl w:val="{k-1}"/>` |
| **参考文献标题** (`参考文献`) | 若存在则设置 | **1 级** (`outline_level = 1`) | `<w:outlineLvl w:val="0"/>` |
| **附录总标题** (`附　　录` / `附录`) | 若存在则设置 | **1 级** (`outline_level = 1`) | `<w:outlineLvl w:val="0"/>` |
| **附录分项标题** (`附录 A ...`、`附录 B ...`、`附录 C ...`) | 若存在则设置 | **2 级** (`outline_level = 2`) | `<w:outlineLvl w:val="1"/>` |
| **附录内子标题** (e.g., `A.1 ...`、`A.1.1 ...`) | 若存在则顺延 | **3 级 / 4 级...** (`outline_level = 3, 4...`) | `<w:outlineLvl w:val="2"/>`, `w:val="3"`... |
| **封面、中英文摘要、目录条目、正文段落、图题、表题、参考文献条目** | 严禁设置大纲级别 | **正文文本** (`outline_level = None`) | 不写入 `<w:outlineLvl>` |

### Step 5: Enforce Section Page Breaks & Native Word Automatic TOC (`自动生成目录与独立换页`, Rules 2.2–2.9)

1. **Mandatory Page Breaks Between Document Sections (`文档主要章节间严格换页`)**:
   - **Cover $\rightarrow$ Chinese Abstract**: Insert a page break (`<w:br w:type="page"/>`) after the cover block (Rule 2.1).
   - **Chinese Abstract $\rightarrow$ English Abstract**: Insert a page break (`<w:br w:type="page"/>`) immediately after the Chinese keywords paragraph (Rules 2.2–2.3) so that Chinese and English abstracts are ALWAYS on separate pages.
   - **Preceding Section $\rightarrow$ Table of Contents (`目录前必有换页符`)**: Insert a page break after the English keywords paragraph (Rules 2.4–2.5). Even if the English abstract block is omitted, the Table of Contents (Rule 2.6) MUST always be preceded by a page break so it starts on a fresh page.
   - **Table of Contents $\rightarrow$ Main Body**: Insert a page break (or Next-Page Section Break) after the Table of Contents so Chapter 1 starts on a fresh page.
   - **Main Body $\rightarrow$ References (`参考文献前必有换页符`)**: Insert a page break before the "参考文献" heading (Rule 2.8) so that references ALWAYS start on a fresh page.
   - **References (or Main Body) $\rightarrow$ Appendices (`附录前必有换页符`)**: Insert a page break before the "附　　录" / "附录" heading (Rule 2.9) so that appendices ALWAYS start on a fresh page.

2. **Native Word Automatic Table of Contents (`Word 原生自动目录域与页码链接`, Rule 2.6)**:
   - **Title & Blank Line**: Centered `目　　　录` (3 full-width spaces between characters, `黑体` `三号 16pt`), followed by 1 blank line in `五号 (10.5pt)`.
   - **Native `TOC` Field (`TOC \o "1-3" \h \z \u`)**:
     - Instead of plain static text without page numbers, construct a genuine Word complex field `TOC \o "1-3" \h \z \u` (where `\u` harvests all paragraphs carrying `<w:outlineLvl>` levels 1–3, and `\h` enables clickable hyperlink navigation).
     - Attach a unique `<w:bookmarkStart w:id="..." w:name="_Toc..."/>` / `<w:bookmarkEnd/>` bookmark to every outline-level heading in the document.
     - Each TOC entry inside the `TOC` field result contains a nested `HYPERLINK \l _Toc...` field, right-aligned dot-leader tab (`<w:tab w:val="right" w:leader="dot" w:pos="9072"/>`), and a nested `PAGEREF _Toc... \h` field displaying the physical page number.
     - When Microsoft Word COM (`win32com.client`) is available on the build host, automatically invoke `Fields.Update()` during build to populate exact physical page numbers (`1, 2, 3...`) and re-apply Rule 2.6 typography (`一级黑体四号顶格`, `二级宋体小四缩进2字符/24pt`, `三级宋体小四缩进4字符/48pt`, ASCII/page numbers in `Times New Roman`). Also set `<w:updateFields w:val="true"/>` in `word/settings.xml`.

3. **Page Numbering & Section Break Specification (`页码格式与正文起始页规范`, Rule 2.17)**:
   - **Front-Matter vs. Main Body Section Break (`前置部分与正文分节`)**:
     - Place all front-matter sections (Cover, Chinese Abstract, English Abstract, Table of Contents) in the Front-Matter Section (`Section 0`), and insert a **Next-Page Section Break (`分节符（下一页）`, `<w:sectPr><w:type w:val="nextPage"/></w:sectPr>`)** immediately after the Table of Contents (before Chapter 1 of the Main Body).
     - **Front-Matter Footer**: Unpopulated/empty centered paragraph (`<w:pPr><w:jc w:val="center"/></w:pPr>`) — front-matter pages (Cover, Abstracts, TOC) do NOT display page numbers.
   - **Main Body Page Numbering Starts at Page 1 (`从正文开始为第 1 页`)**:
     - In the Main Body section (`Section 1`), unlink the footer from the previous section (`footer.is_linked_to_previous = False`) and set `<w:pgNumType w:start="1"/>` on `<w:sectPr>` so that the first page of the Main Body is numbered **`1`** (and Chapter 1 in the automatic TOC displays page **`1`**).
   - **Footer Page Number Format (`页脚页码样式`)**:
     - **Position & Alignment**: Bottom of page, centered (`<w:jc w:val="center"/>`), footer distance from bottom edge `w:footer="992"` dxa (`1.75 cm`), header distance `w:header="851"` dxa (`1.5 cm`).
     - **SDT & Field Structure**: Wrap the footer paragraph in a native Word Page Number building block `<w:sdt>` (`<w:docPartGallery w:val="Page Numbers (Bottom of Page)"/><w:docPartUnique/>`), containing the field `PAGE   \* MERGEFORMAT` (plain Arabic numerals `1`, `2`, `3`..., without dashes or `"第 X 页"` text).
     - **Typography**: `小五号 (9.0 pt)` (`<w:sz w:val="18"/><w:szCs w:val="18"/>`), `Times New Roman` for digits/ASCII and `宋体` for East Asian (`w:lang w:val="zh-CN"`).

### Step 6: Section-by-Section Summary Table (Rules 2.1 to 2.17)

| Rule ID | Document Element | Chinese Font | English/Number Font | Size (`Pt`) | Word Outline Level (`大纲级别`) | Alignment & Special Rules |
| :---: | :--- | :---: | :---: | :---: | :---: | :--- |
| **2.1** | 封面题目与信息 | 题目黑体 / 信息宋体 | Times New Roman | 小二 / 三号 / 小三 | 正文文本 (无) | 居中，1.5倍行距，**末尾必加换页符** |
| **2.2-2.3** | 中文摘要与关键词 | 题目黑体 / 摘要楷体 | Times New Roman | 三号 / 四号 / 五号 / 小四 | 正文文本 (无) | 摘要首行缩进2字符；关键词全角“；”隔开无缩进；**与英文摘要之间必加换页符** |
| **2.4-2.5** | 英文摘要与关键词 | - | Times New Roman | 三号 / 四号 / 五号 / 小四 | 正文文本 (无) | Abstract与Key Words加粗；关键词半角“; ”隔开无缩进；**末尾必加换页符** |
| **2.6** | “目　　录”与自动目录域 | 一级黑体 / 二三级宋体 | Times New Roman | 三号 / 四号 / 小四 | 正文文本 (无) | **与前一部分之间必有换页符**；标题间空3个汉字空格，下空1行五号字；采用 **Word 原生 `TOC \o "1-3" \h \z \u` 自动目录域**（含前导点线、真实页码与点击跳转超链接）；二级缩进2字符，三级缩进4字符；**末尾必加换页符** |
| **2.7** | 正文标题与段落 | 一二级黑体 / 三级与正文宋体 | Times New Roman / Cambria Math | 一级四号 / 其余小四 | **标题 1/2/3/... 级** (`val=0/1/2/...`)；段落为正文文本 | 标题顶格并绑定对应大纲级别；正文首行缩进2字符 (`24pt`)，行内数学符号使用 `<m:oMath>` |
| **2.8/2.14** | 参考文献 | 标题黑体 / 条目宋体 | Times New Roman | 标题四号 / 条目小四 | **标题 1 级** (`val=0`，存在即设)；条目为正文文本 | 标题顶格，条目首行无缩进，符合 GB/T 7714-2015；**与正文最后一章之间必加换页符** |
| **2.9/2.15** | 附录与分项 A,B,C | 标题黑体 / 正文楷体 | Times New Roman | 标题四号 / 分项与正文小四 | **“附录”总标题 1 级** (`val=0`)；**附录 A/B/C 分项 2 级** (`val=1`) | 附录存在即设 1 级大纲级别，附录中 A、B、C 等分项设为 2 级大纲级别；编号图A1、表A1、式(A1)；**与前一部分（参考文献或正文）之间必加换页符** |
| **2.10** | 插图与图题 | 宋体 | Times New Roman | 五号 (`10.5pt`) | 正文文本 (无) | 居中，图题在图正下方，图序与图题间空一格无标点 |
| **2.11** | 简明三线表与表题 | 宋体 | Times New Roman | 五号 (`10.5pt`) | 正文文本 (无) | 表题在表上方正中空一格无标点；顶底线 `1.5pt`、栏目线 `0.75pt`、无竖线 |
| **2.12** | 数学公式与编号 | 宋体 | Cambria Math / Times New Roman | 小四 (`12pt`) / 长公式 11pt | 正文文本 (无) | 原生 `<m:oMath>` 公式居中，编号 `(3-1)` 靠右对齐；中间列宽 13.8cm 零内边距，严禁出现 `⬚` 虚线空框，单行排版 |
| **2.16** | 大纲级别层级体系 | 随所属标题规则 | 随所属标题规则 | 随所属标题规则 | **1~9 级** (`w:val="0"~"8"`) | 一级标题/参考文献/附录为 1 级，二级标题/附录 A,B,C 为 2 级，三级标题/附录子标题为 3 级，依次类推 |
| **2.17** | 页码格式与起始页 | 宋体 | Times New Roman | 小五 (`9pt`, `sz=18`) | 正文文本 (无) | 目录后插入**分节符（下一页）**；前置部分不显示页码；**从正文第 1 章开始为第 1 页**（`<w:pgNumType w:start="1"/>`），页面底端居中纯数字（`PAGE   \* MERGEFORMAT`） |
| **2.18** | 语言与校对元数据 | - | - | - | - | 显式绑定 `w:lang w:val="en-US" w:eastAsia="zh-CN"`，严禁 `ja-JP`；开启 `hideSpellingErrors` 消除专有名词假性红线；默认 `updateFields=false` 避免打开安全弹窗 |

## Gotchas

- **Always use `w:val="nil"` instead of omitting border tags**: Omitting border tags on cells allows default Word table styles to draw unwanted gridlines or drop the header bottom line.
- **Never use plain text underscores for math**: Even simple variables like $x_p$, $I_1$, or $N_a$ must be rendered via `<m:oMath><m:sSub>...</m:sSub></m:oMath>`.
- **Zero empty OMML slots (`empty_math_slots == 0`)**: Ensure all `<m:nary>` operands, fractions `<m:f>`, and `<m:rad>` have valid contents, preventing Word `⬚` dashed placeholder boxes.
- **Remember `<w:outlineLvl>` is 0-indexed in OpenXML**: Level 1 outline is `w:val="0"`, Level 2 outline is `w:val="1"`, Level 3 outline is `w:val="2"`, etc. Always set `"参考文献"` and `"附录"` to Level 1 (`w:val="0"`) when present, and `"附录 A"`, `"附录 B"`, `"附录 C"`, etc. to Level 2 (`w:val="1"`).
