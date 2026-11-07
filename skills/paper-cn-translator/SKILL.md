---
name: paper-cn-translator
description: >-
  Translate English computer vision and AI research papers into rigorous, publication-grade Chinese prose. Supports Dual-Mode execution: Mode A (Multimodal Vision-First for Gemini Flash/GPT-4o reading rendered high-DPI page images) and Mode B (Pure-Text Layout-Aware De-Interleaved Extraction for DeepSeek/text-only models), with native PDF stream & object-level precision figure extraction. Use whenever translating academic papers or PDF manuscripts into Chinese.
---
# Paper CN Translator (Dual-Mode: Vision + Pure-Text Edition)

A dual-mode workflow for translating English academic papers (PDF, HTML, or LaTeX) into publication-grade Chinese prose. It automatically adapts to the capabilities of the executing model and extracts standalone figures directly from the PDF's internal object tree:

- **Mode A (Multimodal Vision Mode - Default for Gemini Flash / Vision LLMs)**: Rasterizes PDF pages into 250-DPI PNG images so the vision encoder reads two-column layouts, complex LaTeX equations, and multi-header tables visually, while **automatically extracting standalone figures** via native PDF image streams (`xref`) and exact PDF graphic-object bounding boxes into `parsed_paper/figures/fig_<id>.png` (never crudely slicing figures out of full-page screenshots).
- **Mode B (Pure-Text Layout-Aware Mode - Default for DeepSeek-V3/R1 or Text-Only LLMs)**: Uses a center-gutter-aware spatial bounding-box de-interleaving algorithm (`pdfplumber` + `PyMuPDF`/`pypdfium2`) that separates full-width headers/tables from left and right columns, extracts structured Markdown tables, and runs the same precision PDF object-level figure extractor so text-only models never suffer from two-column text scrambling or mis-cropped figures.

## When to Use

- The user wants a full, unabridged Chinese translation of an English research paper (e.g. CVPR, ICCV, ECCV, NeurIPS, ICML, ICLR, AAAI).
- The executing agent is either a multimodal vision model (such as Gemini Flash) OR a pure-text model (such as DeepSeek).
- The translated output needs to feed into `cn-thesis-format-spec` and `docx-plugin-engine` to produce a strictly formatted Word (`.docx`) document.

## Steps

### Step 1: Select Execution Mode & Extract Clean Standalone Figures

Check whether your current model environment supports viewing image files (`view_file` / `read_file` multimodal vision input):

1. **If Multimodal Vision is available -> Use Mode A (Vision-First + Object-Level Figure Extraction)**:

```bash
python3 skills/paper-cn-translator/scripts/pdf_vision_preprocessor.py render \
  --pdf input_paper.pdf \
  --outdir parsed_paper \
  --dpi 250
```

   This single command:
   - Renders `parsed_paper/page_01.png` ... `page_N.png` for visual reading of two-column text, equations, and tables in 2-to-3-page visual chunks;
   - Automatically invokes the **PDF Object-Level Figure Extractor (`extract-figures`)** into `parsed_paper/figures/fig_1.png`, `fig_2.png`, ... and writes `parsed_paper/figures/figures_manifest.json`:
     - **Single-Bitmap Figures**: Extracted directly from the native embedded PDF image stream (`fitz.Pixmap(doc, xref)`) at 100% original resolution with zero page background or surrounding text.
     - **Composite / Vector / Multi-Panel Figures**: Computed from the exact mathematical union of PDF graphic primitives (`image_rects` + `vector drawings` + figure-internal axis/legend/subfigure labels) clustered bottom-up above each `Figure X.` caption, excluding body paragraphs, headings, and the English `Figure X.` caption itself, then rendered at 300 DPI and tight-trimmed with a uniform `10 px` white border.
   - You can also run standalone figure extraction at any time via:

```bash
python3 skills/paper-cn-translator/scripts/pdf_vision_preprocessor.py extract-figures \
  --pdf input_paper.pdf \
  --outdir parsed_paper/figures \
  --dpi 300
```

2. **If Running on a Pure-Text Model (or without image viewing) -> Use Mode B (Layout-Aware Text + Object-Level Figure Mode)**:

```bash
python3 skills/paper-cn-translator/scripts/pdf_vision_preprocessor.py extract-text \
  --pdf input_paper.pdf \
  --outdir parsed_paper \
  --dpi 300
```

   This single command automatically:
   - Splits each page into **Top Full-Width Span -> Left Column -> Right Column -> Bottom Full-Width Span** using gutter-aware word `x0/x1` segmentation so left and right columns never interleave;
   - Extracts all PDF tables into Markdown tables inside `parsed_paper/paper_deinterleaved.md`;
   - Runs `extract_pdf_figures` to extract all standalone figures (`parsed_paper/figures/fig_1.png`, `fig_2.png`, ...) via native PDF streams and exact PDF graphic-object bounding boxes, recording their paths and captions in `parsed_paper/extraction_manifest.json`;
   - Reads `parsed_paper/paper_deinterleaved.md` in 4-to-5-page chunks to translate the entire paper without truncation.

### Step 2: Lock Domain Terminology Before Translating

Enforce a unified bilingual glossary across all chunks:

| English Term | Standard Chinese Translation | Notes |
| :--- | :--- | :--- |
| Training-Free | 免训练（Training-Free） | Never translate as 无训练 |
| Few-Shot Anomaly Detection | 少样本异常检测 | Keep $k$-shot as $k$-shot 或 $k$ 样本 |
| Subspace Modeling | 子空间建模 | Linear subspace -> 线性子空间 |
| Vision Foundation Model (VFM) | 视觉基础模型 | DINOv2 / CLIP keep original names |
| Patch Token / Register Token | 图像块令牌 / 寄存器令牌 | Keep Patch in parentheses |
| Principal Component Analysis (PCA) | 主成分分析（PCA） | Eigenvector -> 特征向量 |
| Reconstruction Residual | 重构残差 | Orthogonal projection -> 正交投影 |
| Tail Value-at-Risk (TVaR) | 尾部风险价值（TVaR） | Top 1% patch scores -> 前 1% 块得分 |
| Per-Region Overlap (PRO) | 区域重叠率（PRO） | Pixel-AUROC -> 像素级 AUROC |

### Step 3: Strict Mathematical Notation Rules (Zero Plain-Text Underscores)

Whether reading from page images (Mode A) or de-interleaved text (Mode B), convert every mathematical variable, subscript, superscript, or formula into clean LaTeX syntax:

- **Wrap every inline symbol or formula in `$...$`**: Write `$\mathcal{I}_{\mathrm{train}} = \{I_1, \dots, I_k\}$`, `$I_{\mathrm{test}}$`, `$f_l(p) \in \mathbb{R}^D$`, `$x_p \in \mathbb{R}^D$`, `$N_a = 30$`, `$X_{\mathrm{normal}} \in \mathbb{R}^{n \times D}$`, `$\mu \in \mathbb{R}^D$`, `$\Sigma \in \mathbb{R}^{D \times D}$`, `$\|(x - \mu) - C C^\top (x - \mu)\|_2^2$`, `$H_\rho(M)$`, `$s_{\mathrm{img}}$`, `$O(n D^2)$`, and `$10^6$`.
- **Repair broken PDF math tokens in Mode B**: When pure-text extraction turns a subscript like $x_p$ into `xp` or `x p`, use surrounding context and equation definitions to restore `$x_p$` and `$\sum_{i=1}^{r} \lambda_i \ge \tau \sum_{i=1}^{D} \lambda_i$`.

### Step 4: Emit Structured JSON (`translated_paper_data.json`)

Output the translated content as a structured JSON file (`translated_paper_data.json`) containing `cover`, `cn_abstract_block`, `en_abstract_block`, `toc`, `body_blocks` (with `heading`, `paragraph`, `equation`, `table` with `group_headers` and `__GROUP__` rows, and `figure`), `references`, and `appendices`.

## Gotchas

- **Never summarize paragraphs in either mode**: Translate every sentence completely from Section 1 through the Appendices.
- **Never manually guess full-page crop coordinates for standard figures**: Always use the clean standalone figures generated in `parsed_paper/figures/fig_1.png`, `fig_2.png`, ... (listed in `parsed_paper/figures/figures_manifest.json` and `parsed_paper/extraction_manifest.json`) so figures never contain sliced body text, author headers, or cut-off plot axes.
