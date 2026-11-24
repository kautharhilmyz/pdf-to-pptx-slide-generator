"""Final Verification & Optional Visual QA (PPTX XML check + LibreOffice PNG render)."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
from pathlib import Path

import pymupdf
from pptx import Presentation


def find_soffice() -> str | None:
    """Locate LibreOffice soffice binary if installed."""
    exe = shutil.which("soffice")
    if exe:
        return exe
    win_candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for candidate in win_candidates:
        if candidate.exists():
            return str(candidate)
    return None


def inspect_pptx(pptx_path: Path) -> dict:
    """Verify PPTX opens cleanly and count slides, images, and speaker notes."""
    prs = Presentation(str(pptx_path))
    slides_summary = []
    total_images = 0
    notes_populated = 0

    for idx, slide in enumerate(prs.slides, start=1):
        img_count = sum(1 for shape in slide.shapes if shape.shape_type == 13)  # PICTURE
        total_images += img_count
        notes_text = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
        if notes_text:
            notes_populated += 1
        title_text = slide.shapes.title.text.strip() if slide.shapes.title else ""
        slides_summary.append(
            {
                "slide": idx,
                "title": title_text,
                "images": img_count,
                "has_notes": bool(notes_text),
            }
        )

    return {
        "slide_count": len(prs.slides),
        "total_images": total_images,
        "notes_populated": notes_populated,
        "slides": slides_summary,
    }


def render_previews(pptx_path: Path, render_dir: Path) -> list[str]:
    """Convert PPTX -> PDF via LibreOffice, then PDF -> PNG via PyMuPDF."""
    soffice = find_soffice()
    if not soffice:
        return []

    render_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(render_dir), str(pptx_path)],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    pdf_path = render_dir / f"{pptx_path.stem}.pdf"
    if not pdf_path.exists():
        return []

    png_paths: list[str] = []
    doc = pymupdf.open(pdf_path)
    for idx, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=150)
        png_path = render_dir / f"slide_{idx:02d}.png"
        pix.save(str(png_path))
        png_paths.append(str(png_path))
    doc.close()
    return png_paths


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify PPTX integrity and optionally render PNG previews.")
    parser.add_argument("--pptx", required=True, help="Path to generated .pptx file")
    parser.add_argument("--render-dir", default="", help="Optional directory to render slide PNG previews")
    args = parser.parse_args()

    pptx_path = Path(args.pptx).resolve()
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")

    report = inspect_pptx(pptx_path)
    if args.render_dir:
        previews = render_previews(pptx_path, Path(args.render_dir).resolve())
        report["visual_qa_rendered"] = bool(previews)
        report["preview_files"] = previews
    else:
        report["visual_qa_rendered"] = False

    print(json.dumps({"status": "ok", **report}, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
