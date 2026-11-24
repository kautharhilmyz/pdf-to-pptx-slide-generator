"""Phase 1: Multi-format source extraction for pdf-to-pptx."""

from __future__ import annotations

import argparse
import io
import json
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

import pdfplumber
import pymupdf
from PIL import Image
from pptx import Presentation


DOCX_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def extract_pdf(
    pdf_path: Path,
    file_idx: int,
    assets_dir: Path,
    text_chunks: list[str],
    tables_out: list[dict],
    images_out: list[dict],
) -> None:
    """Extract text, tables, and filtered images from a PDF file."""
    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)
    xref_pages: Counter[int] = Counter()

    for page_index in range(total_pages):
        page = doc[page_index]
        page_num = page_index + 1
        text = page.get_text("text").strip()
        if text:
            text_chunks.append(f"[FILE: {pdf_path.name} | PAGE {page_num}]\n{text}\n")
        for img in page.get_images(full=True):
            xref_pages[img[0]] += 1

    seen_xrefs: set[int] = set()
    for page_index in range(total_pages):
        page = doc[page_index]
        page_num = page_index + 1
        for img_idx, img_info in enumerate(page.get_images(full=True)):
            xref = img_info[0]
            if xref in seen_xrefs:
                continue
            if total_pages > 2 and (xref_pages[xref] / total_pages) > 0.40:
                continue
            seen_xrefs.add(xref)
            try:
                base_img = doc.extract_image(xref)
                img_bytes = base_img["image"]
                pil_img = Image.open(io.BytesIO(img_bytes))
                w, h = pil_img.size
                if w < 200 or h < 200:
                    continue
                aspect = w / max(h, 1)
                if aspect < 0.15 or aspect > 6.5:
                    continue
                if pil_img.mode not in ("RGB", "RGBA"):
                    pil_img = pil_img.convert("RGB")
                out_name = f"pdf_img_f{file_idx}_p{page_num}_{img_idx}.png"
                out_path = assets_dir / out_name
                pil_img.save(out_path, format="PNG")
                images_out.append(
                    {
                        "source_file": pdf_path.name,
                        "page": page_num,
                        "path": f"assets/{out_name}",
                        "width": w,
                        "height": h,
                    }
                )
            except Exception:
                continue
    doc.close()

    with pdfplumber.open(pdf_path) as plumber_pdf:
        for page_index, page in enumerate(plumber_pdf.pages):
            page_num = page_index + 1
            for table in page.extract_tables() or []:
                cleaned = [
                    [cell.strip() if isinstance(cell, str) else "" for cell in row]
                    for row in table
                    if row and any(cell and str(cell).strip() for cell in row)
                ]
                if len(cleaned) < 2:
                    continue
                tables_out.append(
                    {
                        "source_file": pdf_path.name,
                        "page": page_num,
                        "headers": cleaned[0],
                        "rows": cleaned[1:],
                    }
                )


def extract_docx(docx_path: Path, text_chunks: list[str], tables_out: list[dict]) -> None:
    """Extract paragraphs and tables from a DOCX file using built-in zipfile + XML."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            return
        root = ET.fromstring(zf.read("word/document.xml"))

    paragraphs: list[str] = []
    for p in root.findall(".//w:body/w:p", DOCX_NS):
        texts = [t.text for t in p.findall(".//w:t", DOCX_NS) if t.text]
        line = "".join(texts).strip()
        if line:
            paragraphs.append(line)

    if paragraphs:
        text_chunks.append(f"[FILE: {docx_path.name} | DOCX]\n" + "\n".join(paragraphs) + "\n")

    for tbl in root.findall(".//w:body/w:tbl", DOCX_NS):
        rows: list[list[str]] = []
        for tr in tbl.findall("./w:tr", DOCX_NS):
            row = []
            for tc in tr.findall("./w:tc", DOCX_NS):
                cell_texts = [t.text for t in tc.findall(".//w:t", DOCX_NS) if t.text]
                row.append("".join(cell_texts).strip())
            if any(row):
                rows.append(row)
        if len(rows) >= 2:
            tables_out.append(
                {
                    "source_file": docx_path.name,
                    "page": 1,
                    "headers": rows[0],
                    "rows": rows[1:],
                }
            )


def extract_text_file(txt_path: Path, text_chunks: list[str]) -> None:
    """Extract UTF-8 text from .txt or .md files."""
    content = txt_path.read_text(encoding="utf-8", errors="replace").strip()
    if content:
        text_chunks.append(f"[FILE: {txt_path.name} | TEXT]\n{content}\n")


def inspect_sample_pptx(sample_paths: list[Path]) -> list[dict]:
    """Inspect reference PPTX files for dimensions, fonts, colors, and layouts."""
    results: list[dict] = []
    for path in sample_paths:
        if not path.exists():
            continue
        prs = Presentation(str(path))
        fonts: Counter[str] = Counter()
        colors: Counter[str] = Counter()
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            fonts[run.font.name] += 1
                        try:
                            if run.font.color and run.font.color.rgb:
                                colors[f"#{run.font.color.rgb}"] += 1
                        except AttributeError:
                            pass
        results.append(
            {
                "file": path.name,
                "width_inches": round(prs.slide_width.inches, 3),
                "height_inches": round(prs.slide_height.inches, 3),
                "slide_count": len(prs.slides),
                "layouts": [layout.name for layout in prs.slide_layouts],
                "top_fonts": [f for f, _ in fonts.most_common(5)],
                "top_colors": [c for c, _ in colors.most_common(6)],
            }
        )
    return results


def resolve_input(deck_dir: Path, item: str) -> Path:
    """Resolve a filename relative to deck_dir/source or deck_dir."""
    p = Path(item)
    if p.is_absolute() and p.exists():
        return p
    if (deck_dir / item).exists():
        return deck_dir / item
    return deck_dir / "source" / item


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract text, tables, and images into deck workspace.")
    parser.add_argument("--deck-dir", required=True, help="Target deck directory")
    parser.add_argument("--sources", nargs="+", required=True, help="Source files (.pdf, .docx, .txt, .md)")
    parser.add_argument("--samples", nargs="*", default=[], help="Optional reference .pptx files")
    args = parser.parse_args()

    deck_dir = Path(args.deck_dir).resolve()
    workspace_dir = deck_dir / "workspace"
    assets_dir = deck_dir / "assets"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    assets_dir.mkdir(parents=True, exist_ok=True)

    text_chunks: list[str] = []
    tables_out: list[dict] = []
    images_out: list[dict] = []

    for idx, src in enumerate(args.sources):
        src_path = resolve_input(deck_dir, src)
        if not src_path.exists():
            raise FileNotFoundError(f"Source file not found: {src_path}")
        suffix = src_path.suffix.lower()
        if suffix == ".pdf":
            extract_pdf(src_path, idx, assets_dir, text_chunks, tables_out, images_out)
        elif suffix == ".docx":
            extract_docx(src_path, text_chunks, tables_out)
        elif suffix in (".txt", ".md"):
            extract_text_file(src_path, text_chunks)
        else:
            raise ValueError(f"Unsupported source format: {suffix}")

    (workspace_dir / "raw_text.md").write_text("\n---\n\n".join(text_chunks), encoding="utf-8")
    (workspace_dir / "tables.json").write_text(json.dumps(tables_out, indent=2, ensure_ascii=False), encoding="utf-8")
    (workspace_dir / "image_manifest.json").write_text(
        json.dumps(images_out, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    if args.samples:
        sample_paths = [resolve_input(deck_dir, s) for s in args.samples]
        sample_styles = inspect_sample_pptx(sample_paths)
        (workspace_dir / "sample_style.json").write_text(
            json.dumps(sample_styles, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    print(
        json.dumps(
            {
                "status": "ok",
                "text_sections": len(text_chunks),
                "tables": len(tables_out),
                "images": len(images_out),
            }
        )
    )


if __name__ == "__main__":
    main()
