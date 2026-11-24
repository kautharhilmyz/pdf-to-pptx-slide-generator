"""Phase 3: Data-driven 16:9 executive PowerPoint builder for pdf-to-pptx."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


THEMES_PATH = Path(__file__).resolve().parent.parent / "references" / "themes.json"
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN_X = Inches(0.65)
CONTENT_W = SLIDE_W - (MARGIN_X * 2)


def hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert #RRGGBB hex string to python-pptx RGBColor."""
    clean = hex_str.lstrip("#")
    return RGBColor(int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16))


def load_theme(theme_key: str, brand_path: Path | None) -> dict:
    """Load base theme and merge optional brand.json overrides."""
    themes = json.loads(THEMES_PATH.read_text(encoding="utf-8"))
    theme = dict(themes.get(theme_key, themes["classic_navy_teal"]))
    if brand_path and brand_path.exists():
        brand = json.loads(brand_path.read_text(encoding="utf-8"))
        theme.update({k: v for k, v in brand.items() if v is not None})
    return theme


def set_slide_bg(slide, color_hex: str) -> None:
    """Apply solid background color to a slide."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(color_hex)


def add_rect(slide, left, top, width, height, fill_hex: str, border_hex: str | None = None, rounded: bool = False):
    """Add a clean flat rectangle or rounded rectangle."""
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if border_hex:
        shape.line.color.rgb = hex_to_rgb(border_hex)
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    if rounded and shape.adjustments:
        shape.adjustments[0] = 0.06
    return shape


def add_header(slide, title: str, theme: dict, deck_dir: Path) -> None:
    """Render consistent top accent bar, concise slide title, and optional brand logo."""
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.08), theme["primary"])
    tb = slide.shapes.add_textbox(MARGIN_X, Inches(0.32), CONTENT_W - Inches(1.4), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.name = theme.get("font_title", "Segoe UI")
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["dark"])

    add_rect(slide, MARGIN_X, Inches(1.02), Inches(1.1), Inches(0.045), theme["accent"])

    logo_rel = theme.get("logo_path")
    if logo_rel:
        logo_path = Path(logo_rel) if Path(logo_rel).is_absolute() else deck_dir / logo_rel
        if logo_path.exists():
            place_image_fit(slide, logo_path, SLIDE_W - MARGIN_X - Inches(1.1), Inches(0.24), Inches(1.1), Inches(0.55))


def add_bullets(tf, bullets: list[str], theme: dict, font_size: int = 15) -> None:
    """Populate a text frame with bold lead-in bullets and 1.2x line spacing."""
    tf.word_wrap = True
    for idx, raw in enumerate(bullets):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        p.line_spacing = 1.2
        match = re.match(r"^\*\*(.+?)\*\*\s*(.*)$", raw.strip())
        if match:
            lead, rest = match.group(1), match.group(2)
            r1 = p.add_run()
            r1.text = f"▪  {lead} "
            r1.font.name = theme.get("font_body", "Calibri")
            r1.font.size = Pt(font_size)
            r1.font.bold = True
            r1.font.color.rgb = hex_to_rgb(theme["primary"])
            if rest:
                r2 = p.add_run()
                r2.text = rest
                r2.font.name = theme.get("font_body", "Calibri")
                r2.font.size = Pt(font_size)
                r2.font.color.rgb = hex_to_rgb(theme["body"])
        else:
            r = p.add_run()
            r.text = f"▪  {raw}"
            r.font.name = theme.get("font_body", "Calibri")
            r.font.size = Pt(font_size)
            r.font.color.rgb = hex_to_rgb(theme["body"])


def place_image_fit(slide, img_path: Path, box_left, box_top, box_w, box_h) -> None:
    """Place image centered inside target bounding box while preserving exact aspect ratio."""
    with Image.open(img_path) as im:
        img_w, img_h = im.size
    if img_w <= 0 or img_h <= 0:
        return
    scale = min(box_w / img_w, box_h / img_h)
    final_w = int(img_w * scale)
    final_h = int(img_h * scale)
    left = int(box_left + (box_w - final_w) / 2)
    top = int(box_top + (box_h - final_h) / 2)
    slide.shapes.add_picture(str(img_path), left, top, width=final_w, height=final_h)


def build_title_slide(slide, spec: dict, theme: dict, deck_dir: Path) -> None:
    set_slide_bg(slide, theme["primary"])
    add_rect(slide, Inches(0), Inches(0), Inches(0.28), SLIDE_H, theme["accent"])

    tb = slide.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(11.2), Inches(3.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p_title = tf.paragraphs[0]
    r_title = p_title.add_run()
    r_title.text = spec.get("title", "Presentation")
    r_title.font.name = theme.get("font_title", "Segoe UI")
    r_title.font.size = Pt(40)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(255, 255, 255)

    if spec.get("subtitle"):
        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(14)
        r_sub = p_sub.add_run()
        r_sub.text = spec["subtitle"]
        r_sub.font.name = theme.get("font_body", "Calibri")
        r_sub.font.size = Pt(20)
        r_sub.font.color.rgb = hex_to_rgb("#E2E8F0")

    if spec.get("meta"):
        p_meta = tf.add_paragraph()
        p_meta.space_before = Pt(28)
        r_meta = p_meta.add_run()
        r_meta.text = spec["meta"]
        r_meta.font.name = theme.get("font_body", "Calibri")
        r_meta.font.size = Pt(14)
        r_meta.font.color.rgb = hex_to_rgb("#CBD5E1")

    logo_rel = theme.get("logo_path")
    if logo_rel:
        logo_path = Path(logo_rel) if Path(logo_rel).is_absolute() else deck_dir / logo_rel
        if logo_path.exists():
            place_image_fit(slide, logo_path, SLIDE_W - Inches(2.2), Inches(0.5), Inches(1.5), Inches(0.75))


def build_section_header(slide, spec: dict, theme: dict) -> None:
    set_slide_bg(slide, theme["dark"])
    add_rect(slide, Inches(0), SLIDE_H - Inches(0.14), SLIDE_W, Inches(0.14), theme["accent"])

    tb = slide.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(11.0), Inches(2.8))
    tf = tb.text_frame
    tf.word_wrap = True
    if spec.get("section_number"):
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = spec["section_number"]
        r0.font.name = theme.get("font_title", "Segoe UI")
        r0.font.size = Pt(22)
        r0.font.bold = True
        r0.font.color.rgb = hex_to_rgb(theme["accent"])
        p1 = tf.add_paragraph()
        p1.space_before = Pt(8)
    else:
        p1 = tf.paragraphs[0]

    r1 = p1.add_run()
    r1.text = spec.get("title", "")
    r1.font.name = theme.get("font_title", "Segoe UI")
    r1.font.size = Pt(36)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(255, 255, 255)

    if spec.get("subtitle"):
        p2 = tf.add_paragraph()
        p2.space_before = Pt(12)
        r2 = p2.add_run()
        r2.text = spec["subtitle"]
        r2.font.name = theme.get("font_body", "Calibri")
        r2.font.size = Pt(18)
        r2.font.color.rgb = hex_to_rgb("#CBD5E1")


def build_two_column_visual(slide, spec: dict, theme: dict, deck_dir: Path) -> None:
    set_slide_bg(slide, theme["background"])
    add_header(slide, spec.get("title", ""), theme, deck_dir)

    tb = slide.shapes.add_textbox(MARGIN_X, Inches(1.35), Inches(5.3), Inches(5.5))
    add_bullets(tb.text_frame, spec.get("bullets", []), theme, font_size=15)

    vis_rel = spec.get("visual_path")
    if vis_rel:
        vis_path = Path(vis_rel) if Path(vis_rel).is_absolute() else deck_dir / vis_rel
        if vis_path.exists():
            add_rect(slide, Inches(6.2), Inches(1.3), Inches(6.48), Inches(5.6), theme["surface"], theme["border"], rounded=True)
            place_image_fit(slide, vis_path, Inches(6.35), Inches(1.45), Inches(6.18), Inches(5.3))


def build_key_metrics_cards(slide, spec: dict, theme: dict, deck_dir: Path) -> None:
    set_slide_bg(slide, theme["background"])
    add_header(slide, spec.get("title", ""), theme, deck_dir)

    cards = spec.get("cards", [])[:4]
    if cards:
        gap = Inches(0.28)
        card_w = (CONTENT_W - gap * (len(cards) - 1)) / len(cards)
        for i, card in enumerate(cards):
            left = MARGIN_X + i * (card_w + gap)
            add_rect(slide, left, Inches(1.35), card_w, Inches(2.15), theme["surface"], theme["border"], rounded=True)
            add_rect(slide, left, Inches(1.35), card_w, Inches(0.08), theme["accent"] if i % 2 else theme["primary"])

            tb = slide.shapes.add_textbox(left + Inches(0.22), Inches(1.55), card_w - Inches(0.44), Inches(1.8))
            tf = tb.text_frame
            tf.word_wrap = True
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = card.get("label", "")
            r0.font.name = theme.get("font_body", "Calibri")
            r0.font.size = Pt(13)
            r0.font.bold = True
            r0.font.color.rgb = hex_to_rgb(theme["muted"])

            p1 = tf.add_paragraph()
            p1.space_before = Pt(8)
            r1 = p1.add_run()
            r1.text = card.get("value", "")
            r1.font.name = theme.get("font_title", "Segoe UI")
            r1.font.size = Pt(30)
            r1.font.bold = True
            r1.font.color.rgb = hex_to_rgb(theme["primary"])

            if card.get("delta"):
                p2 = tf.add_paragraph()
                p2.space_before = Pt(6)
                r2 = p2.add_run()
                r2.text = card["delta"]
                r2.font.name = theme.get("font_body", "Calibri")
                r2.font.size = Pt(12)
                r2.font.bold = True
                r2.font.color.rgb = hex_to_rgb(theme["accent"])

    if spec.get("bullets"):
        tb = slide.shapes.add_textbox(MARGIN_X, Inches(3.85), CONTENT_W, Inches(3.0))
        add_bullets(tb.text_frame, spec["bullets"], theme, font_size=15)


def build_comparison_table(slide, spec: dict, theme: dict, deck_dir: Path) -> None:
    set_slide_bg(slide, theme["background"])
    add_header(slide, spec.get("title", ""), theme, deck_dir)

    headers = spec.get("headers", [])
    rows = spec.get("rows", [])
    if not headers:
        return
    n_rows = len(rows) + 1
    n_cols = len(headers)
    table_shape = slide.shapes.add_table(n_rows, n_cols, MARGIN_X, Inches(1.4), CONTENT_W, Inches(min(5.4, 0.62 * n_rows)))
    table = table_shape.table

    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = hex_to_rgb(theme["primary"])
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = str(h_text)
        r.font.name = theme.get("font_title", "Segoe UI")
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, row_data in enumerate(rows, start=1):
        bg_hex = theme["surface"] if r_idx % 2 == 1 else "#F1F5F9"
        for c_idx in range(n_cols):
            val = row_data[c_idx] if c_idx < len(row_data) else ""
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb(bg_hex)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = str(val)
            r.font.name = theme.get("font_body", "Calibri")
            r.font.size = Pt(13)
            r.font.bold = c_idx == 0
            r.font.color.rgb = hex_to_rgb(theme["dark"] if c_idx == 0 else theme["body"])


def build_executive_summary(slide, spec: dict, theme: dict, deck_dir: Path) -> None:
    set_slide_bg(slide, theme["background"])
    add_header(slide, spec.get("title", "Executive Summary"), theme, deck_dir)

    cards = spec.get("cards", [])[:4]
    if cards:
        gap = Inches(0.28)
        card_w = (CONTENT_W - gap * (len(cards) - 1)) / len(cards)
        for i, card in enumerate(cards):
            left = MARGIN_X + i * (card_w + gap)
            add_rect(slide, left, Inches(1.38), card_w, Inches(3.55), theme["surface"], theme["border"], rounded=True)
            add_rect(slide, left, Inches(1.38), card_w, Inches(0.09), theme["primary"] if i % 2 == 0 else theme["accent"])

            tb = slide.shapes.add_textbox(left + Inches(0.22), Inches(1.62), card_w - Inches(0.44), Inches(3.15))
            tf = tb.text_frame
            tf.word_wrap = True
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = card.get("title", "")
            r0.font.name = theme.get("font_title", "Segoe UI")
            r0.font.size = Pt(17)
            r0.font.bold = True
            r0.font.color.rgb = hex_to_rgb(theme["primary"])

            p1 = tf.add_paragraph()
            p1.space_before = Pt(10)
            p1.line_spacing = 1.2
            r1 = p1.add_run()
            r1.text = card.get("body", "")
            r1.font.name = theme.get("font_body", "Calibri")
            r1.font.size = Pt(14)
            r1.font.color.rgb = hex_to_rgb(theme["body"])

    if spec.get("takeaway"):
        add_rect(slide, MARGIN_X, Inches(5.25), CONTENT_W, Inches(1.55), theme["primary"], rounded=True)
        tb = slide.shapes.add_textbox(MARGIN_X + Inches(0.3), Inches(5.42), CONTENT_W - Inches(0.6), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = "KEY TAKEAWAY"
        r0.font.name = theme.get("font_title", "Segoe UI")
        r0.font.size = Pt(11)
        r0.font.bold = True
        r0.font.color.rgb = hex_to_rgb("#93C5FD")

        p1 = tf.add_paragraph()
        p1.space_before = Pt(6)
        r1 = p1.add_run()
        r1.text = spec["takeaway"]
        r1.font.name = theme.get("font_body", "Calibri")
        r1.font.size = Pt(16)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(255, 255, 255)


def build_timeline_roadmap(slide, spec: dict, theme: dict, deck_dir: Path) -> None:
    set_slide_bg(slide, theme["background"])
    add_header(slide, spec.get("title", "Roadmap"), theme, deck_dir)

    milestones = spec.get("milestones", [])[:5]
    if not milestones:
        return
    add_rect(slide, MARGIN_X + Inches(0.5), Inches(2.55), CONTENT_W - Inches(1.0), Inches(0.06), theme["border"])

    gap = Inches(0.24)
    col_w = (CONTENT_W - gap * (len(milestones) - 1)) / len(milestones)
    for i, ms in enumerate(milestones):
        left = MARGIN_X + i * (col_w + gap)
        badge_color = theme["primary"] if i % 2 == 0 else theme["accent"]

        badge = add_rect(slide, left, Inches(1.75), col_w, Inches(0.52), badge_color, rounded=True)
        p_b = badge.text_frame.paragraphs[0]
        p_b.alignment = PP_ALIGN.CENTER
        rb = p_b.add_run()
        rb.text = ms.get("period", f"Phase {i + 1}")
        rb.font.name = theme.get("font_title", "Segoe UI")
        rb.font.size = Pt(13)
        rb.font.bold = True
        rb.font.color.rgb = RGBColor(255, 255, 255)

        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + col_w / 2 - Inches(0.14), Inches(2.44), Inches(0.28), Inches(0.28))
        dot.fill.solid()
        dot.fill.fore_color.rgb = hex_to_rgb(badge_color)
        dot.line.color.rgb = RGBColor(255, 255, 255)
        dot.line.width = Pt(2)

        add_rect(slide, left, Inches(2.95), col_w, Inches(3.6), theme["surface"], theme["border"], rounded=True)
        tb = slide.shapes.add_textbox(left + Inches(0.18), Inches(3.15), col_w - Inches(0.36), Inches(3.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = ms.get("title", "")
        r0.font.name = theme.get("font_title", "Segoe UI")
        r0.font.size = Pt(16)
        r0.font.bold = True
        r0.font.color.rgb = hex_to_rgb(theme["dark"])

        p1 = tf.add_paragraph()
        p1.space_before = Pt(10)
        p1.line_spacing = 1.2
        r1 = p1.add_run()
        r1.text = ms.get("detail", "")
        r1.font.name = theme.get("font_body", "Calibri")
        r1.font.size = Pt(13)
        r1.font.color.rgb = hex_to_rgb(theme["body"])


def build_process_flow(slide, spec: dict, theme: dict, deck_dir: Path) -> None:
    set_slide_bg(slide, theme["background"])
    add_header(slide, spec.get("title", "Process Flow"), theme, deck_dir)

    steps = spec.get("steps", [])[:5]
    if not steps:
        return
    gap = Inches(0.26)
    col_w = (CONTENT_W - gap * (len(steps) - 1)) / len(steps)
    for i, st in enumerate(steps):
        left = MARGIN_X + i * (col_w + gap)
        hdr_color = theme["primary"] if i % 2 == 0 else theme["accent"]

        add_rect(slide, left, Inches(1.65), col_w, Inches(4.9), theme["surface"], theme["border"], rounded=True)
        hdr = add_rect(slide, left, Inches(1.65), col_w, Inches(0.75), hdr_color)
        p_h = hdr.text_frame.paragraphs[0]
        p_h.alignment = PP_ALIGN.CENTER
        rh = p_h.add_run()
        rh.text = f"STEP {st.get('step', i + 1)}"
        rh.font.name = theme.get("font_title", "Segoe UI")
        rh.font.size = Pt(14)
        rh.font.bold = True
        rh.font.color.rgb = RGBColor(255, 255, 255)

        tb = slide.shapes.add_textbox(left + Inches(0.2), Inches(2.6), col_w - Inches(0.4), Inches(3.7))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = st.get("title", "")
        r0.font.name = theme.get("font_title", "Segoe UI")
        r0.font.size = Pt(17)
        r0.font.bold = True
        r0.font.color.rgb = hex_to_rgb(theme["dark"])

        p1 = tf.add_paragraph()
        p1.space_before = Pt(10)
        p1.line_spacing = 1.2
        r1 = p1.add_run()
        r1.text = st.get("detail", "")
        r1.font.name = theme.get("font_body", "Calibri")
        r1.font.size = Pt(14)
        r1.font.color.rgb = hex_to_rgb(theme["body"])


def build_matrix_2x2(slide, spec: dict, theme: dict, deck_dir: Path) -> None:
    set_slide_bg(slide, theme["background"])
    add_header(slide, spec.get("title", "2x2 Matrix"), theme, deck_dir)

    quads = spec.get("quadrants", {})
    grid_left = MARGIN_X + Inches(0.75)
    grid_top = Inches(1.35)
    cell_w = (CONTENT_W - Inches(0.95)) / 2
    cell_h = Inches(2.62)
    gap = Inches(0.2)

    positions = [
        ("top_left", grid_left, grid_top, theme["accent"]),
        ("top_right", grid_left + cell_w + gap, grid_top, theme["primary"]),
        ("bottom_left", grid_left, grid_top + cell_h + gap, theme["muted"]),
        ("bottom_right", grid_left + cell_w + gap, grid_top + cell_h + gap, theme["highlight"]),
    ]
    for key, left, top, bar_color in positions:
        q = quads.get(key, {})
        add_rect(slide, left, top, cell_w, cell_h, theme["surface"], theme["border"], rounded=True)
        add_rect(slide, left, top, Inches(0.1), cell_h, bar_color)

        tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.18), cell_w - Inches(0.4), cell_h - Inches(0.32))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = q.get("title", key.replace("_", " ").title())
        r0.font.name = theme.get("font_title", "Segoe UI")
        r0.font.size = Pt(16)
        r0.font.bold = True
        r0.font.color.rgb = hex_to_rgb(bar_color)

        for item in q.get("items", []):
            p = tf.add_paragraph()
            p.space_before = Pt(6)
            r = p.add_run()
            r.text = f"▪  {item}"
            r.font.name = theme.get("font_body", "Calibri")
            r.font.size = Pt(13)
            r.font.color.rgb = hex_to_rgb(theme["body"])

    if spec.get("x_axis"):
        tb_x = slide.shapes.add_textbox(grid_left, Inches(6.9), CONTENT_W - Inches(0.75), Inches(0.35))
        px = tb_x.text_frame.paragraphs[0]
        px.alignment = PP_ALIGN.CENTER
        rx = px.add_run()
        rx.text = spec["x_axis"]
        rx.font.name = theme.get("font_body", "Calibri")
        rx.font.size = Pt(12)
        rx.font.bold = True
        rx.font.color.rgb = hex_to_rgb(theme["muted"])

    if spec.get("y_axis"):
        tb_y = slide.shapes.add_textbox(MARGIN_X - Inches(0.15), Inches(3.8), Inches(0.85), Inches(0.6))
        tb_y.text_frame.word_wrap = True
        py = tb_y.text_frame.paragraphs[0]
        py.alignment = PP_ALIGN.CENTER
        ry = py.add_run()
        ry.text = spec["y_axis"]
        ry.font.name = theme.get("font_body", "Calibri")
        ry.font.size = Pt(11)
        ry.font.bold = True
        ry.font.color.rgb = hex_to_rgb(theme["muted"])


LAYOUT_BUILDERS = {
    "title_slide": lambda s, sp, th, dd: build_title_slide(s, sp, th, dd),
    "section_header": lambda s, sp, th, dd: build_section_header(s, sp, th),
    "two_column_chart": build_two_column_visual,
    "two_column_image": build_two_column_visual,
    "key_metrics_cards": build_key_metrics_cards,
    "comparison_table": build_comparison_table,
    "executive_summary": build_executive_summary,
    "timeline_roadmap": build_timeline_roadmap,
    "process_flow": build_process_flow,
    "matrix_2x2": build_matrix_2x2,
}


def main() -> None:
    parser = argparse.ArgumentParser(description="Build 16:9 PPTX from slide_plan JSON.")
    parser.add_argument("--deck-dir", required=True, help="Target deck directory")
    parser.add_argument("--plan", required=True, help="Path to slide_plan JSON")
    parser.add_argument("--output", required=True, help="Output .pptx path (relative to deck-dir or absolute)")
    args = parser.parse_args()

    deck_dir = Path(args.deck_dir).resolve()
    plan_path = Path(args.plan) if Path(args.plan).is_absolute() else deck_dir / args.plan
    plan = json.loads(plan_path.read_text(encoding="utf-8"))

    brand_file = plan.get("brand_file", "brand.json")
    brand_path = deck_dir / brand_file if brand_file else None
    theme = load_theme(plan.get("theme", "classic_navy_teal"), brand_path)

    template_rel = plan.get("template_pptx")
    if template_rel:
        tpl_path = Path(template_rel) if Path(template_rel).is_absolute() else deck_dir / template_rel
        prs = Presentation(str(tpl_path)) if tpl_path.exists() else Presentation()
    else:
        prs = Presentation()
        prs.slide_width = SLIDE_W
        prs.slide_height = SLIDE_H

    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]

    for spec in plan.get("slides", []):
        slide = prs.slides.add_slide(blank_layout)
        layout_name = spec.get("layout", "executive_summary")
        builder = LAYOUT_BUILDERS.get(layout_name, build_executive_summary)
        builder(slide, spec, theme, deck_dir)

        notes = spec.get("speaker_notes", "").strip()
        if notes:
            slide.notes_slide.notes_text_frame.text = notes

    out_path = Path(args.output) if Path(args.output).is_absolute() else deck_dir / args.output
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    print(json.dumps({"status": "ok", "slides": len(prs.slides), "output": str(out_path)}))


if __name__ == "__main__":
    main()
