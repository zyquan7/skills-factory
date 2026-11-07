#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dual-Mode Academic PDF Preprocessor & Object-Level Figure Extractor for `paper-cn-translator`.
Supports:
1. Precision PDF Figure Extraction (`extract-figures`, also auto-run in `render` and `extract-text`):
   - Inspects the PDF internal object tree (via PyMuPDF `fitz`, with `pdfplumber`+`pypdfium2` object-level fallback).
   - For single-bitmap figures (e.g. Figure 1), extracts the original native embedded image stream (`xref`) directly
     from the PDF without page cropping or resampling.
   - For composite/vector/multi-panel figures (e.g. architecture diagrams, matplotlib/pgfplots vector charts,
     multi-subfigure grids), computes the exact mathematical union bounding box of all PDF graphic primitives
     (embedded image rects + vector drawing paths + figure-internal axis/legend/subfigure labels) via bottom-up
     vertical contiguity clustering while strictly excluding prose paragraphs, headings, page headers, and the
     `Figure X.` caption itself.
2. Mode A (Vision Models like Gemini Flash):
   - `render`: Rasterizes PDF pages into 250-DPI PNG images (`page_01.png`, ...) AND automatically extracts
     all standalone figures into `<outdir>/figures/fig_<id>.png`.
   - `crop`: Crops a region from a page PNG using normalized [ymin, xmin, ymax, xmax] (0..1000) coordinates
     with automatic whitespace trimming.
3. Mode B (Pure-Text Models like DeepSeek-V3/R1):
   - `extract-text`: Performs layout-aware two-column de-interleaved text extraction with center-gutter splitting,
     extracts structured Markdown tables, and runs precision object-level figure extraction.
"""

import os
import re
import json
import shutil
import argparse
from typing import Dict, List, Optional, Tuple

import pypdfium2 as pdfium
import pdfplumber
from PIL import Image

try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False


CAPTION_START_RE = re.compile(r"^(?:Figure|Fig\.)\s*(\d+|[A-Z]\d*)\s*[\.:]", re.IGNORECASE)
NON_FIG_BLOCK_RE = re.compile(
    r"^(?:Table\s+\d+|Figure\s+\d+|Fig\.\s*\d+|Abstract\b|References\b|\d+(?:\.\d+)*\.?\s+[A-Z][a-zA-Z]{2,})"
)


def _tight_trim_whitespace(img: Image.Image, threshold: int = 248, pad_px: int = 10) -> Image.Image:
    """
    Trims outer near-white margins using pure Pillow C-level LUT + getbbox()
    and adds a clean, uniform `pad_px` white border.
    """
    rgb = img.convert("RGB")
    gray = rgb.convert("L")
    mask = gray.point(lambda p: 255 if p < threshold else 0)
    bbox = mask.getbbox()
    if not bbox:
        return rgb

    x0, y0, x1, y1 = bbox
    cropped = rgb.crop((
        max(0, x0),
        max(0, y0),
        min(rgb.width, x1),
        min(rgb.height, y1),
    ))
    if pad_px <= 0:
        return cropped

    canvas = Image.new("RGB", (cropped.width + 2 * pad_px, cropped.height + 2 * pad_px), (255, 255, 255))
    canvas.paste(cropped, (pad_px, pad_px))
    return canvas


def _extract_figures_fitz(pdf_path: str, fig_dir: str, dpi: int = 300) -> List[Dict]:
    """
    Precision PDF Figure Extractor powered by PyMuPDF (`fitz`):
    - Detects true `Figure X.` / `Fig. X:` caption blocks (`page.get_text("blocks")`).
    - Determines column horizontal bounds (single-column vs full-width spanning figure).
    - Computes hard ceiling `ceiling_y` from preceding prose paragraphs, section headings, tables, or title headers.
    - Clusters PDF graphical primitives (`image_rects` + `drawings` + short figure-internal labels) bottom-up from
      the caption top (`max_gap_pt = 22.0`) so unrelated elements above the figure are never merged.
    - Extracts native embedded image stream (`fitz.Pixmap(doc, xref)`) when the figure is a single standalone bitmap,
      or renders the exact PDF object-union bounding box at `dpi` with tight whitespace trimming for vector/composite figures.
    """
    os.makedirs(fig_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    extracted_figures: List[Dict] = []

    for pno in range(len(doc)):
        page = doc[pno]
        page_num = pno + 1
        W, H = float(page.rect.width), float(page.rect.height)
        mid_x = W / 2.0
        blocks = page.get_text("blocks")

        # 1. Identify all true figure caption blocks on this page
        captions = []
        for b in blocks:
            txt = b[4].strip()
            m = CAPTION_START_RE.match(txt)
            if m:
                clean_cap = " ".join(line.strip() for line in txt.splitlines() if line.strip())
                captions.append((m.group(1), fitz.Rect(b[:4]), clean_cap))

        if not captions:
            continue

        # Sort captions top-to-bottom, left-to-right
        captions.sort(key=lambda item: (item[1].y0, item[1].x0))

        # 2. Gather embedded raster image rects on this page (excluding page header/footer watermarks)
        img_items: List[Tuple[int, fitz.Rect]] = []
        for im in page.get_images(full=True):
            xref = im[0]
            for r in page.get_image_rects(xref):
                if r.y1 > 42.0 and r.y0 < H - 52.0 and r.width >= 4.0 and r.height >= 4.0:
                    img_items.append((xref, fitz.Rect(r)))

        # 3. Gather vector drawing rects on this page (excluding header/footer watermarks and sub-point dust)
        draw_rects: List[fitz.Rect] = []
        for d in page.get_drawings():
            r = d.get("rect")
            if not r:
                continue
            fr = fitz.Rect(r)
            if fr.y1 <= 42.0 or fr.y0 >= H - 52.0:
                continue
            if fr.width < 0.5 and fr.height < 0.5:
                continue
            draw_rects.append(fr)

        # 4. For each caption, resolve the exact figure object cluster above it
        for local_idx, (fig_id, cap_rect, cap_txt) in enumerate(captions, start=1):
            spans_cols = (cap_rect.x0 < mid_x - W * 0.05) and (cap_rect.x1 > mid_x + W * 0.05)
            if spans_cols:
                col_xmin, col_xmax = 0.0, W
            elif cap_rect.x1 <= mid_x + W * 0.05:
                col_xmin, col_xmax = 0.0, mid_x + 12.0
            else:
                col_xmin, col_xmax = mid_x - 12.0, W

            # Find ceiling_y from prose paragraphs, headings, tables, or earlier captions above cap_rect.y0
            ceiling_y = 42.0
            for b in blocks:
                br = fitz.Rect(b[:4])
                btxt = b[4].strip()
                if not btxt or br == cap_rect:
                    continue
                if br.x1 < col_xmin + 10.0 or br.x0 > col_xmax - 10.0:
                    continue
                if br.y1 <= cap_rect.y0 + 1.0:
                    lines = [ln.strip() for ln in btxt.splitlines() if ln.strip()]
                    avg_line_len = len(btxt) / max(1, len(lines))
                    is_prose_or_hdr = (
                        NON_FIG_BLOCK_RE.match(btxt) is not None
                        or (len(btxt) >= 150 and avg_line_len >= 35.0)
                        or (pno == 0 and br.y1 < 220.0 and br.x0 < mid_x and br.x1 > mid_x)
                    )
                    if is_prose_or_hdr:
                        ceiling_y = max(ceiling_y, br.y1 + 2.0)

            cand_imgs = [
                (xref, r) for xref, r in img_items
                if r.x1 > col_xmin and r.x0 < col_xmax and r.y0 >= ceiling_y - 4.0 and r.y1 <= cap_rect.y0 + 2.0
            ]
            cand_draws = [
                r for r in draw_rects
                if r.x1 > col_xmin and r.x0 < col_xmax and r.y0 >= ceiling_y - 4.0 and r.y1 <= cap_rect.y0 + 2.0
            ]
            cand_txts: List[fitz.Rect] = []
            for b in blocks:
                br = fitz.Rect(b[:4])
                btxt = b[4].strip()
                if not btxt or br == cap_rect:
                    continue
                if (
                    br.x1 > col_xmin
                    and br.x0 < col_xmax
                    and br.y0 >= ceiling_y - 2.0
                    and br.y1 <= cap_rect.y0 - 0.5
                    and len(btxt) < 150
                    and not NON_FIG_BLOCK_RE.match(btxt)
                ):
                    cand_txts.append(br)

            # Bottom-up vertical contiguity clustering from cap_rect.y0 upward (stops at vertical gaps > 22pt)
            tagged_objs: List[Tuple[str, fitz.Rect, Optional[int]]] = []
            for xref, r in cand_imgs:
                tagged_objs.append(("img", r, xref))
            for r in cand_draws:
                tagged_objs.append(("draw", r, None))
            for r in cand_txts:
                tagged_objs.append(("txt", r, None))

            if not tagged_objs:
                continue

            tagged_objs.sort(key=lambda item: item[1].y1, reverse=True)
            cluster_objs: List[Tuple[str, fitz.Rect, Optional[int]]] = []
            cur_top_y = cap_rect.y0
            max_vertical_gap_pt = 22.0

            for kind, r, xref in tagged_objs:
                gap = cur_top_y - r.y1
                if cluster_objs and gap > max_vertical_gap_pt:
                    break
                cluster_objs.append((kind, r, xref))
                cur_top_y = min(cur_top_y, r.y0)

            cluster_imgs = [(xref, r) for kind, r, xref in cluster_objs if kind == "img" and xref is not None]
            cluster_draws = [r for kind, r, _ in cluster_objs if kind == "draw"]
            cluster_txts = [r for kind, r, _ in cluster_objs if kind == "txt"]

            if not cluster_imgs and not cluster_draws:
                continue

            primary_filename = f"fig_{fig_id}.png"
            primary_path = os.path.join(fig_dir, primary_filename)
            alias_filename = f"fig_page{page_num:02d}_{local_idx}.png"
            alias_path = os.path.join(fig_dir, alias_filename)

            # Case 1: Single native embedded bitmap with no vector drawings or external text labels
            if len(cluster_imgs) == 1 and len(cluster_draws) == 0 and len(cluster_txts) == 0:
                xref, img_rect = cluster_imgs[0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                pix.save(primary_path)
                if os.path.abspath(primary_path) != os.path.abspath(alias_path):
                    shutil.copyfile(primary_path, alias_path)
                fig_w, fig_h = pix.width, pix.height
                union_r = img_rect
                mode_str = "native_embedded_stream"
            else:
                # Case 2: Exact PDF object-union bounding box rendered at high DPI + tight whitespace trim
                all_rects = [r for _, r, _ in cluster_objs]
                union_r = fitz.Rect(all_rects[0])
                for r in all_rects[1:]:
                    union_r |= r

                clip_r = fitz.Rect(
                    max(0.0, union_r.x0 - 4.0),
                    max(ceiling_y, union_r.y0 - 4.0),
                    min(W, union_r.x1 + 4.0),
                    min(cap_rect.y0 - 1.5, union_r.y1 + 3.0),
                )
                pix = page.get_pixmap(clip=clip_r, dpi=dpi, alpha=False)
                pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                pil_img = _tight_trim_whitespace(pil_img, threshold=248, pad_px=10)
                pil_img.save(primary_path, "PNG")
                if os.path.abspath(primary_path) != os.path.abspath(alias_path):
                    shutil.copyfile(primary_path, alias_path)
                fig_w, fig_h = pil_img.width, pil_img.height
                mode_str = "pdf_object_union_render"

            info = {
                "figure_id": f"Figure {fig_id}",
                "page": page_num,
                "extraction_mode": mode_str,
                "caption_text": cap_txt,
                "image_path": primary_path,
                "cropped_image_path": primary_path,
                "page_alias_path": alias_path,
                "width": fig_w,
                "height": fig_h,
                "bbox_pt": [round(union_r.x0, 2), round(union_r.y0, 2), round(union_r.x1, 2), round(union_r.y1, 2)],
            }
            extracted_figures.append(info)
            print(
                f"[Figure Extractor] Page {page_num:02d} | Figure {fig_id} ({mode_str}) "
                f"-> {primary_path} ({fig_w}x{fig_h} px)"
            )

    doc.close()
    return extracted_figures


def _extract_figures_plumber_fallback(pdf_path: str, fig_dir: str, dpi: int = 300) -> List[Dict]:
    """
    Object-level fallback figure extractor using `pdfplumber` + `pypdfium2` if PyMuPDF (`fitz`) is unavailable.
    Computes the union bounding box of `page.images`, `page.rects`, `page.curves`, and `page.lines`
    clustered above each `Figure X.` caption.
    """
    os.makedirs(fig_dir, exist_ok=True)
    pdfium_doc = pdfium.PdfDocument(pdf_path)
    scale = dpi / 72.0
    extracted_figures: List[Dict] = []

    with pdfplumber.open(pdf_path) as plumber_pdf:
        for idx, page in enumerate(plumber_pdf.pages):
            page_num = idx + 1
            W, H = float(page.width), float(page.height)
            mid_x = W / 2.0
            words = page.extract_words(x_tolerance=2.0, y_tolerance=3.0)
            lines = _group_words_into_lines(words, y_tol=3.0, max_x_gap=W * 0.028)

            captions = []
            for ln in lines:
                txt = ln["text"].strip()
                m = CAPTION_START_RE.match(txt)
                if m:
                    captions.append((m.group(1), ln))

            if not captions:
                continue

            pil_page = pdfium_doc[idx].render(scale=scale).to_pil().convert("RGB")
            PW, PH = pil_page.size

            gfx_boxes = []
            for obj_list in (page.images, page.rects, page.curves, page.lines):
                for obj in obj_list:
                    x0, top, x1, bottom = float(obj["x0"]), float(obj["top"]), float(obj["x1"]), float(obj["bottom"])
                    if bottom > 42.0 and top < H - 52.0 and (x1 - x0 >= 0.5 or bottom - top >= 0.5):
                        gfx_boxes.append((x0, top, x1, bottom))

            for local_idx, (fig_id, cap) in enumerate(captions, start=1):
                spans_cols = (cap["x0"] < mid_x - W * 0.05) and (cap["x1"] > mid_x + W * 0.05)
                col_xmin, col_xmax = (0.0, W) if spans_cols else (
                    (0.0, mid_x + 12.0) if cap["x1"] <= mid_x + W * 0.05 else (mid_x - 12.0, W)
                )
                cands = [
                    b for b in gfx_boxes
                    if b[2] > col_xmin and b[0] < col_xmax and 42.0 <= b[1] and b[3] <= cap["top"] + 2.0
                ]
                if not cands:
                    continue
                cands.sort(key=lambda b: b[3], reverse=True)
                cluster = []
                cur_top = float(cap["top"])
                for b in cands:
                    if cluster and (cur_top - b[3]) > 22.0:
                        break
                    cluster.append(b)
                    cur_top = min(cur_top, b[1])

                ux0 = max(0.0, min(b[0] for b in cluster) - 6.0)
                uy0 = max(42.0, min(b[1] for b in cluster) - 16.0)
                ux1 = min(W, max(b[2] for b in cluster) + 6.0)
                uy1 = min(float(cap["top"]) - 1.5, max(b[3] for b in cluster) + 14.0)

                crop_px = (
                    int(ux0 * PW / W),
                    int(uy0 * PH / H),
                    int(ux1 * PW / W),
                    int(uy1 * PH / H),
                )
                cropped = _tight_trim_whitespace(pil_page.crop(crop_px), threshold=248, pad_px=10)
                primary_path = os.path.join(fig_dir, f"fig_{fig_id}.png")
                alias_path = os.path.join(fig_dir, f"fig_page{page_num:02d}_{local_idx}.png")
                cropped.save(primary_path, "PNG")
                if os.path.abspath(primary_path) != os.path.abspath(alias_path):
                    shutil.copyfile(primary_path, alias_path)

                extracted_figures.append({
                    "figure_id": f"Figure {fig_id}",
                    "page": page_num,
                    "extraction_mode": "pdfplumber_object_union_render",
                    "caption_text": cap["text"],
                    "image_path": primary_path,
                    "cropped_image_path": primary_path,
                    "page_alias_path": alias_path,
                    "width": cropped.width,
                    "height": cropped.height,
                    "bbox_pt": [round(ux0, 2), round(uy0, 2), round(ux1, 2), round(uy1, 2)],
                })

    return extracted_figures


def extract_pdf_figures(pdf_path: str, fig_dir: str, dpi: int = 300) -> List[Dict]:
    """
    Extracts all standalone figures from `pdf_path` into `fig_dir` using PDF internal object coordinates
    and native image streams, and writes `figures_manifest.json`.
    """
    if HAS_FITZ:
        figs = _extract_figures_fitz(pdf_path, fig_dir, dpi=dpi)
    else:
        figs = _extract_figures_plumber_fallback(pdf_path, fig_dir, dpi=dpi)

    manifest_path = os.path.join(fig_dir, "figures_manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump({"pdf_path": pdf_path, "figures": figs}, f, indent=2, ensure_ascii=False)
    print(f"[Figure Extractor] Extracted {len(figs)} figures cleanly -> {fig_dir} (manifest: {manifest_path})")
    return figs


def render_pdf_pages(pdf_path: str, outdir: str, dpi: int = 250, extract_figs: bool = True) -> List[str]:
    """
    Mode A (Vision):
    1. Renders full PDF pages into `page_01.png` ... `page_N.png` for visual reading of text/equations/tables.
    2. Automatically extracts clean standalone figures into `<outdir>/figures/fig_X.png` via `extract_pdf_figures`
       so the vision model never has to crudely crop figures out of full-page screenshots.
    """
    os.makedirs(outdir, exist_ok=True)
    doc = pdfium.PdfDocument(pdf_path)
    scale = dpi / 72.0
    saved_paths = []
    print(f"[Preprocessor Mode A] Rendering {len(doc)} pages from {pdf_path} at {dpi} DPI...")
    for idx in range(len(doc)):
        page = doc[idx]
        bitmap = page.render(scale=scale)
        pil_img = bitmap.to_pil()
        out_p = os.path.join(outdir, f"page_{idx + 1:02d}.png")
        pil_img.save(out_p, "PNG")
        saved_paths.append(out_p)
        print(f"  -> Saved {out_p} ({pil_img.width}x{pil_img.height} px)")

    if extract_figs:
        fig_dir = os.path.join(outdir, "figures")
        extract_pdf_figures(pdf_path, fig_dir, dpi=max(300, dpi))

    return saved_paths


def crop_normalized_bbox(image_path: str, bbox: List[int], out_path: str, pad_px: int = 6, auto_trim: bool = True) -> None:
    """
    Crops a normalized [ymin, xmin, ymax, xmax] (0..1000) region from `image_path`
    and applies automatic whitespace trimming.
    """
    img = Image.open(image_path).convert("RGB")
    W, H = img.size
    ymin, xmin, ymax, xmax = bbox
    left = max(0, int(round(xmin * W / 1000.0)) - pad_px)
    top = max(0, int(round(ymin * H / 1000.0)) - pad_px)
    right = min(W, int(round(xmax * W / 1000.0)) + pad_px)
    bottom = min(H, int(round(ymax * H / 1000.0)) + pad_px)

    cropped = img.crop((left, top, right, bottom))
    if auto_trim:
        cropped = _tight_trim_whitespace(cropped, threshold=248, pad_px=10)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    cropped.save(out_path, "PNG")
    print(f"[Preprocessor] Cropped bbox {bbox} -> {out_path} ({cropped.width}x{cropped.height} px)")


def _group_words_into_lines(
    words: List[Dict], y_tol: float = 3.5, max_x_gap: Optional[float] = None
) -> List[Dict]:
    """
    Groups pdfplumber words into horizontal lines sorted by vertical top coordinate.
    When `max_x_gap` is specified, splits horizontal segments separated by a wide column gutter
    so left-column and right-column lines at the same vertical position never merge into one line.
    """
    if not words:
        return []
    sorted_w = sorted(words, key=lambda w: (round(w["top"] / y_tol), w["x0"]))
    row_buckets: List[List[Dict]] = []
    cur_bucket = [sorted_w[0]]
    for w in sorted_w[1:]:
        if abs(w["top"] - cur_bucket[-1]["top"]) <= y_tol:
            cur_bucket.append(w)
        else:
            row_buckets.append(cur_bucket)
            cur_bucket = [w]
    row_buckets.append(cur_bucket)

    lines: List[Dict] = []
    for bucket in row_buckets:
        bucket.sort(key=lambda x: x["x0"])
        seg = [bucket[0]]
        for w in bucket[1:]:
            if max_x_gap is not None and (w["x0"] - seg[-1]["x1"]) > max_x_gap:
                lines.append({
                    "top": min(x["top"] for x in seg),
                    "bottom": max(x["bottom"] for x in seg),
                    "x0": min(x["x0"] for x in seg),
                    "x1": max(x["x1"] for x in seg),
                    "text": " ".join(x["text"] for x in seg),
                })
                seg = [w]
            else:
                seg.append(w)
        lines.append({
            "top": min(x["top"] for x in seg),
            "bottom": max(x["bottom"] for x in seg),
            "x0": min(x["x0"] for x in seg),
            "x1": max(x["x1"] for x in seg),
            "text": " ".join(x["text"] for x in seg),
        })
    return sorted(lines, key=lambda l: (l["top"], l["x0"]))


def deinterleave_two_column_page(page) -> Tuple[str, List[Dict]]:
    """
    Mode B Core Algorithm:
    Separates words on a PDF page into:
    1. Top Full-Width Header/Title/Figure-Caption lines (genuinely spanning across the center gutter)
    2. Left Column lines
    3. Right Column lines
    4. Bottom Full-Width Table/Footnote lines
    """
    W = float(page.width)
    H = float(page.height)
    mid_x = W / 2.0

    words = page.extract_words(x_tolerance=2.0, y_tolerance=3.0, keep_blank_chars=False)
    if not words:
        return "", []

    # Use gutter-aware line grouping so left and right columns at the same y never merge
    raw_lines = _group_words_into_lines(words, y_tol=3.0, max_x_gap=W * 0.028)

    top_span_lines = []
    bottom_span_lines = []
    left_lines = []
    right_lines = []
    figure_captions = []

    for ln in raw_lines:
        txt = ln["text"].strip()
        if CAPTION_START_RE.match(txt):
            figure_captions.append(ln)

        crosses_center = (ln["x0"] < mid_x - W * 0.06) and (ln["x1"] > mid_x + W * 0.06)
        if crosses_center:
            if ln["top"] < H * 0.45:
                top_span_lines.append(txt)
            else:
                bottom_span_lines.append(txt)
        else:
            line_center = 0.5 * (ln["x0"] + ln["x1"])
            if line_center < mid_x:
                left_lines.append(txt)
            else:
                right_lines.append(txt)

    sections = []
    if top_span_lines:
        sections.append("### [Full-Width Top Block]\n" + "\n".join(top_span_lines))
    if left_lines:
        sections.append("### [Left Column]\n" + "\n".join(left_lines))
    if right_lines:
        sections.append("### [Right Column]\n" + "\n".join(right_lines))
    if bottom_span_lines:
        sections.append("### [Full-Width Bottom Block]\n" + "\n".join(bottom_span_lines))

    return "\n\n".join(sections), figure_captions


def extract_text_and_auto_figures(pdf_path: str, outdir: str, dpi: int = 300) -> None:
    """
    Mode B Full Pipeline for Pure-Text Models:
    1. Extracts all figures cleanly via `extract_pdf_figures` (using native PDF streams and exact PDF
       object-union bounding boxes, never blind page-ratio cropping).
    2. Extracts layout-aware de-interleaved text and structured Markdown tables per page via `pdfplumber`.
    3. Saves `paper_deinterleaved.md` and `extraction_manifest.json`.
    """
    os.makedirs(outdir, exist_ok=True)
    fig_dir = os.path.join(outdir, "figures")
    os.makedirs(fig_dir, exist_ok=True)

    all_figures = extract_pdf_figures(pdf_path, fig_dir, dpi=dpi)
    figs_by_page: Dict[int, List[Dict]] = {}
    for f_info in all_figures:
        figs_by_page.setdefault(f_info["page"], []).append(f_info)

    md_chunks = [f"# De-Interleaved Academic Paper Text & Tables ({os.path.basename(pdf_path)})\n"]
    manifest = {
        "pdf_path": pdf_path,
        "pages": [],
        "extracted_figures": all_figures,
        "auto_cropped_figures": all_figures,
    }

    with pdfplumber.open(pdf_path) as plumber_pdf:
        for idx, page in enumerate(plumber_pdf.pages):
            page_num = idx + 1
            deinterleaved_text, _ = deinterleave_two_column_page(page)
            page_figs = figs_by_page.get(page_num, [])

            tables = page.extract_tables()
            table_md_list = []
            for t_i, tbl in enumerate(tables, start=1):
                clean_rows = [
                    [str(cell).replace("\n", " ").strip() if cell else "" for cell in row]
                    for row in tbl if any(row)
                ]
                if len(clean_rows) >= 2:
                    hdr = "| " + " | ".join(clean_rows[0]) + " |"
                    sep = "| " + " | ".join(["---"] * len(clean_rows[0])) + " |"
                    body = "\n".join("| " + " | ".join(r) + " |" for r in clean_rows[1:])
                    table_md_list.append(f"#### [Extracted Table {t_i} on Page {page_num}]\n{hdr}\n{sep}\n{body}")

            page_md = [f"## ==================== Page {page_num} ====================\n", deinterleaved_text]
            if page_figs:
                fig_lines = [
                    f"- **{f['figure_id']}** (`{f['image_path']}`, mode=`{f['extraction_mode']}`, "
                    f"{f['width']}x{f['height']} px): {f['caption_text']}"
                    for f in page_figs
                ]
                page_md.append("### [Extracted Standalone Figures on Page]\n" + "\n".join(fig_lines))
            if table_md_list:
                page_md.append("\n\n".join(table_md_list))

            md_chunks.append("\n\n".join(page_md))
            manifest["pages"].append({
                "page": page_num,
                "num_tables_extracted": len(table_md_list),
                "num_figures_cropped": len(page_figs),
            })

    md_out = os.path.join(outdir, "paper_deinterleaved.md")
    with open(md_out, "w", encoding="utf-8") as f:
        f.write("\n\n".join(md_chunks) + "\n")

    manifest_out = os.path.join(outdir, "extraction_manifest.json")
    with open(manifest_out, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    print(f"[Preprocessor Mode B] Saved de-interleaved text & tables to: {md_out}")
    print(f"[Preprocessor Mode B] Extracted {len(all_figures)} clean figures into: {fig_dir}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Dual-Mode PDF Preprocessor & Object-Level Figure Extractor (Vision Mode A & Pure-Text Mode B)"
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_ren = sub.add_parser(
        "render",
        help="Mode A (Vision): Render PDF pages to high-DPI PNGs AND extract standalone figures into <outdir>/figures",
    )
    p_ren.add_argument("--pdf", type=str, required=True, help="Input PDF path")
    p_ren.add_argument("--outdir", type=str, required=True, help="Output directory for page PNGs and figures/")
    p_ren.add_argument("--dpi", type=int, default=250, help="Render DPI for page PNGs (default: 250)")
    p_ren.add_argument("--no-extract-figures", action="store_true", help="Skip automatic figure extraction during render")

    p_fig = sub.add_parser(
        "extract-figures",
        help="Mode A & B: Extract standalone figures via native PDF streams & exact PDF graphic object bounds",
    )
    p_fig.add_argument("--pdf", type=str, required=True, help="Input PDF path")
    p_fig.add_argument("--outdir", type=str, required=True, help="Output directory for extracted figures")
    p_fig.add_argument("--dpi", type=int, default=300, help="Render DPI for vector/composite figures (default: 300)")

    p_crop = sub.add_parser(
        "crop",
        help="Fallback manual crop from page PNG using [ymin, xmin, ymax, xmax] (0..1000) with auto whitespace trim",
    )
    p_crop.add_argument("--image", type=str, required=True, help="Input page PNG path")
    p_crop.add_argument("--bbox", type=int, nargs=4, required=True, metavar=("YMIN", "XMIN", "YMAX", "XMAX"))
    p_crop.add_argument("--out", type=str, required=True, help="Output cropped figure PNG path")

    p_txt = sub.add_parser(
        "extract-text",
        help="Mode B (Pure-Text): De-interleave 2-column text, extract tables, and extract clean standalone figures",
    )
    p_txt.add_argument("--pdf", type=str, required=True, help="Input PDF path")
    p_txt.add_argument("--outdir", type=str, required=True, help="Output directory for de-interleaved Markdown and figures")
    p_txt.add_argument("--dpi", type=int, default=300, help="DPI for vector/composite figure extraction (default: 300)")

    args = parser.parse_args()
    if args.cmd == "render":
        render_pdf_pages(args.pdf, args.outdir, dpi=args.dpi, extract_figs=not args.no_extract_figures)
    elif args.cmd == "extract-figures":
        extract_pdf_figures(args.pdf, args.outdir, dpi=args.dpi)
    elif args.cmd == "crop":
        crop_normalized_bbox(args.image, args.bbox, args.out)
    elif args.cmd == "extract-text":
        extract_text_and_auto_figures(args.pdf, args.outdir, dpi=args.dpi)
