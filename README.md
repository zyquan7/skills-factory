# Skills Factory

个人自用 AI Agent 技能库（Skill Repository），主要用于集中开发、调试、本地测试与维护个人定制的 Agent 技能，并通过 Git/GitHub 进行纯净版本管理。

---

## 包含的技能清单 (Skills)

### 1. [paper-cn-translator](skills/paper-cn-translator/) — 英文学术论文双模态中翻技能
- **核心定位**：将计算机视觉（CV）、人工智能（AI）领域的顶级学术论文（CVPR、NeurIPS、ICLR、IEEE TPAMI 等）翻译为符合出版级水准的严谨中文科技文献。
- **特色能力**：
  - **双模态自适应**：支持多模态视觉模式（Gemini Flash 等直接阅读 250-DPI 分割页面图）与纯文本模式（DeepSeek 等通过双栏感知分栏防错位算法反交错提取）。
  - **PDF 原生级图表抽取**：利用底层图形树与包围盒聚类算法，自动将单图原生位图流（`xref`）及复合矢量图表按分辨率完整提取，彻底告别截图截取导致的边缘模糊或图文混杂。
  - **无缝对接下游排版**：产出的结构化译文与图表清单可直接送入 `cn-thesis-format-spec` 与 `docx-plugin-engine` 生成 Word 文档。

### 2. [cn-thesis-format-spec](skills/cn-thesis-format-spec/) — 中文论文与科技报告排版规范
- **核心定位**：针对 Word（`.docx`）生成制定的中文学术论文与科技报告严格排版法典。
- **特色能力**：
  - **16 条排版法典（Rules 2.1 ~ 2.16）**：严格定义宋体/黑体/楷体与 Times New Roman 中西文混排字体绑定规则。
  - **精准字号映射**：将中文传统字号（小二号 18pt、三号 16pt、小三号 15pt、四号 14pt、小四号 12pt、五号 10.5pt）准确映射为 OpenXML 半磅值。
  - **学术标准版式**：规范学术三线表（顶线 1.5pt、表头底线 0.75pt、底线 1.5pt）、原生 Word OMML（`<m:oMath>`）行内与行间公式、GB/T 7714 参考文献与标准大纲级别。

### 3. [docx-plugin-engine](skills/docx-plugin-engine/) — OpenXML 原生 Word 文档生成引擎
- **核心定位**：通过 Python (`python-docx`) 直接操作 OpenXML 底层节点，将 Markdown 或结构化内容编译为高保真、零瑕疵的学术级 `.docx` 文档。
- **特色能力**：
  - **原生 Word OMML 公式**：不使用易失效的第三方插件或失真的图片公式，直接写入原生 `<m:oMath>` 节点，支持括号组 `<m:d>`、分式 `<m:f>`、积分求和 `<m:nary>`、根式 `<m:rad>` 等，杜绝转义字符与乱码。
  - **学术三线表防折行机制**：显式计算并注入 `<w:tblGrid>` 与列宽 DXA 绝对值，配置 `<w:noWrap/>`，确保方法名称与数据表头永不出现竖向截断换行。
  - **大纲级别与原生目录**：自动在段落属性中写入 `<w:outlineLvl>`（1~9 级），并配合分页符与标准目录引导符构建 Word 原生大纲结构。

---

## 目录结构

```text
skills_factory/
├── skills/                    # 技能集合目录
│   ├── paper-cn-translator/   # 英文论文双模态中翻技能
│   │   ├── SKILL.md
│   │   └── scripts/
│   │       └── pdf_vision_preprocessor.py
│   ├── cn-thesis-format-spec/ # 中文论文与科技报告排版规范
│   │   └── SKILL.md
│   └── docx-plugin-engine/    # OpenXML 原生 Word 文档生成引擎
│       ├── SKILL.md
│       └── scripts/
│           └── build_formatted_thesis_docx.py
├── .gitignore                 # Git 忽略配置（严格白名单）
└── README.md                  # 仓库说明文档
```
