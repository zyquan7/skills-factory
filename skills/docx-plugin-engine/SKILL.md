---
name: docx-plugin-engine
description: Build and style Microsoft Word (.docx) documents for both Multimodal Vision models (Gemini Flash) and Pure-Text models (DeepSeek) via open-source MCP Word plugins or the bundled deterministic python-docx OpenXML rendering engine with native OMML (<m:oMath>) equations, Word Outline Levels (<w:outlineLvl>), three-line tables, visual QA rendering, and programmatic XML auditing. Use when generating or verifying strictly formatted Word (.docx) documents.
---
# DOCX Plugin Engine (Dual-Mode: Vision QA + Programmatic Text Audit)

A complete toolkit and plugin integration guide enabling both **Multimodal Vision LLMs** (such as Gemini Flash) and **Pure-Text LLMs** (such as DeepSeek-V3/R1) to deterministically produce complex Microsoft Word (`.docx`) documents with native Word MathType/OMML equations (`<m:oMath>`), native Word paragraph outline levels (`<w:outlineLvl>`), and standard Chinese three-line tables (`简明三线表`).

## When to Use

- Configuring Gemini Flash, DeepSeek, or local coding agents with plugins/skills to output `.docx` files.
- Compiling structured paper translations or reports into `.docx` files that strictly obey Chinese thesis formatting standards (`cn-thesis-format-spec`, Rules 2.1 to 2.16).

## Outline Level (`大纲级别`) Construction Support (Rule 2.16)

`scripts/build_formatted_thesis_docx.py` provides both automatic schema-driven outline levels and programmatic API helpers:

1. **Programmatic Helper API (`set_outline_level` & `format_paragraph`)**:
   - `set_outline_level(paragraph, outline_level)`: Writes `<w:pPr><w:outlineLvl w:val="{outline_level - 1}"/></w:pPr>` (where `outline_level=1` maps to OpenXML `w:val="0"` = 1 级大纲级别, `outline_level=2` maps to `w:val="1"` = 2 级大纲级别, ..., up to `9`). Passing `None` leaves the paragraph as `正文文本` (Body Text).
   - `format_paragraph(p, ..., outline_level=None)`: Accepts `outline_level` directly when formatting any paragraph.

2. **Schema-Driven Automatic Outline Level Mapping**:
   - **Body Headings (`body_blocks` with `"type": "heading"`)**: Automatically sets `outline_level = block.get("outline_level", block.get("level", 1))`. Level 1 heading $\rightarrow$ 1 级大纲 (`w:val="0"`), Level 2 heading $\rightarrow$ 2 级大纲 (`w:val="1"`), Level 3 heading $\rightarrow$ 3 级大纲 (`w:val="2"`), and so on for any level $k \in [1, 9]$.
   - **References (`references`)**: When `references` is non-empty, the `"参考文献"` heading automatically receives **1 级大纲级别** (`outline_level = 1`, `w:val="0"`).
   - **Appendices (`appendices`)**: When `appendices` is non-empty:
     - Emits the top-level `"附　　　录"` heading (customizable via `"appendix_section_title"`) with **1 级大纲级别** (`outline_level = 1`, `w:val="0"`);
     - Each appendix item (`"code": "A"`, `"B"`, `"C"`, ...) receives **2 级大纲级别** (`outline_level = app.get("outline_level", 2)`, `w:val="1"`);
     - Sub-headings inside an appendix (`"type": "heading"` inside `app["blocks"]`) automatically shift to **3 级、4 级...大纲级别** (`w:val="2"`, `w:val="3"`, ...), or use an explicit `"outline_level"` if provided.

3. **Section Page Breaks & Native Word Automatic TOC (`自动目录与独立换页`, Rules 2.2–2.9)**:
   - **Guaranteed Page Breaks**: Automatically inserts page breaks (`<w:br w:type="page"/>`) between:
     1. Cover (`cover`) $\rightarrow$ Chinese Abstract (`cn_abstract_block`);
     2. Chinese Abstract (`cn_abstract_block`) $\rightarrow$ English Abstract (`en_abstract_block`);
     3. Preceding section (`en_abstract_block` or `cn_abstract_block`) $\rightarrow$ Table of Contents (`toc`);
     4. Table of Contents (`toc`) $\rightarrow$ Main Body (`body_blocks`);
     5. Main Body (`body_blocks`) $\rightarrow$ References (`references`) (Rule 2.8);
     6. References (`references`) or Main Body $\rightarrow$ Appendices (`appendices`) (Rule 2.9).
   - **Native Word Automatic TOC (`TOC \o "1-3" \h \z \u`) & Zero-Prompt Settings**:
     - Attaches `_Toc...` bookmarks to all outline-level headings and builds a genuine Word complex field `TOC \o "1-3" \h \z \u` containing nested `HYPERLINK \l _Toc...` jump links, right-aligned dot-leader tabs (`9072 dxa`), and `PAGEREF _Toc... \h` page number fields.
     - By default sets `<w:updateFields w:val="false"/>` and strips `w:dirty` flags from fields to prevent Word from prompting "This document contains fields that may refer to other files". Can be toggled via `auto_update_fields: true`.
     - Automatically invokes headless Microsoft Word COM (`win32com.client`) when available to run `Repaginate()` + `Fields.Update()` so that exact physical page numbers are populated into the `.docx` at build time, while preserving Rule 2.6 typography (`一级黑体四号顶格`, `二级宋体小四缩进2字符`, `三级宋体小四缩进4字符`, ASCII/page numbers in `Times New Roman`).
   - **Page Numbering & Section Break Specification (`从正文开始为第 1 页`, Rule 2.17)**:
     - Separates Front-Matter (Cover, Abstracts, TOC) from Main Body using a **Next-Page Section Break (`WD_SECTION.NEW_PAGE`)** after the TOC.
     - Front-Matter section has an empty footer (no page numbers shown on Cover, Abstracts, or TOC).
     - Main Body section unlinks its footer from the previous section (`is_linked_to_previous = False`), sets `<w:pgNumType w:start="1"/>` so **Chapter 1 starts at Page 1** (and shows as page `1` in the automatic TOC), and renders a centered bottom-of-page `<w:sdt>` (`Page Numbers (Bottom of Page)`) containing `PAGE   \* MERGEFORMAT` in `9pt` (`小五号`, `w:sz="18"`, `Times New Roman` / `宋体`).

4. **Robust OMML Math Equations & Single-Line Layout (Rule 2.12)**:
   - **Zero Empty Slot Guarantee**: Recursive argument group parsing for compound commands (`\frac`, nested `\sum`, `\sqrt`, accents `\tilde`, `\hat`, `\bar`) ensures no empty operands `<m:e>`, `<m:num>`, or `<m:den>` are emitted, completely eliminating Word `⬚` dashed placeholder boxes.
   - **Single-Line Equation Table**: Center equation cell expanded to `13.8 cm` with zero cell margins (`tcMar`), combined with equation length sensing and adaptive scaling to prevent awkward equation line wraps.

5. **East Asian Typography & Spell-Check Suppression (Rule 2.18)**:
   - Binds `w:lang w:val="en-US" w:eastAsia="zh-CN"` on all text runs and theme settings, correcting default `ja-JP` misconfigurations that cause massive spell-check red squiggly underlines.
   - Enforces `<w:hideSpellingErrors/>` and `<w:hideGrammaticalErrors/>` in document settings.

## Dual-Mode Verification Workflow (Supports Both Vision & Pure-Text Models)

Run `scripts/build_formatted_thesis_docx.py` with the verification flag suited to your model:

1. **For Multimodal Vision Models (Mode A -> `--render-qa-dir`)**:

```bash
python3 skills/docx-plugin-engine/scripts/build_formatted_thesis_docx.py   --input translated_paper_data.json   --output output_translated_paper.docx   --render-qa-dir qa_previews
```

   Renders 150-DPI page PNGs (`qa_previews/qa_page-01.png`, ...) via LibreOffice + `pdftoppm` so the vision model can visually inspect equations and tables.

2. **For Pure-Text Models (Mode B -> `--verify-xml-audit`)**:

```bash
python3 skills/docx-plugin-engine/scripts/build_formatted_thesis_docx.py   --input translated_paper_data.json   --output output_translated_paper.docx   --verify-xml-audit
```

   Runs a deterministic OpenXML AST audit that prints a text report verifying:
   - Total count of native `<m:oMath>` equation nodes and confirms **zero raw plain-text underscore math tokens** exist in paragraphs, plus **zero empty OMML slots** (`empty_math_slots_count == 0` for `<m:num>`, `<m:den>`, `<m:e>`);
   - Total count of three-line tables and confirms every table has `w:tblW = 9072 dxa` (`16.0 cm`), `w:tblLayout = fixed`, `1.5pt` top/bottom borders (`sz=12`), `0.75pt` header border (`sz=6`), and `nil` vertical borders;
   - **Outline Level Hierarchy (`outline_tree` & `outline_level_counts`)**: Verifies that all body headings ($1, 2, 3, \dots$级), `"参考文献"` (1 级, if present), `"附　　录"` (1 级, if present), and `"附录 A/B/C..."` (2 级, if present) carry exact `<w:outlineLvl>` tags;
   - **Section Page Breaks & Native Automatic TOC (`document_page_breaks` & `native_auto_toc`)**: Verifies explicit page breaks between Chinese and English abstracts, before/after the TOC, between main body and references, and between references and appendices; verifies the `TOC \o "1-3" \h \z \u` field, `HYPERLINK \l _Toc...` jump links, and `PAGEREF _Toc... \h` resolved page numbers;
   - Confirms all East Asian (`宋体`/`黑体`/`楷体`) and ASCII (`Times New Roman`/`Cambria Math`) font bindings, language settings (`zh-CN` without `ja-JP`), and zero-prompt updateFields settings.

## Open-Source Word MCP Servers (For MCP-Enabled Clients)

- **GongRzhe/Office-Word-MCP-Server** (`https://github.com/gongrzhe/office-word-mcp-server`): `pip install office-word-mcp-server`.
- **MeterLong/MCP-Doc** (`https://github.com/meterlong/mcp-doc`): FastMCP-based Word service.
- **anthropics/skills/docx** (`https://github.com/anthropics/skills/blob/main/skills/docx/SKILL.md`): Official Agent Skill (`npx skills add https://github.com/anthropics/skills --skill docx`).

## Gotchas

- **Always run `--verify-xml-audit` on pure-text models**: Even when the model cannot view rendered PNGs, `--verify-xml-audit` guarantees 100% mathematical, outline-level, and structural compliance with Rules 2.1 to 2.16.
