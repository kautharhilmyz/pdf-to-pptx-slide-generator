# AI Slide Generator (`pdf-to-pptx` Skill)

*(Bahasa Melayu: lihat [README-MY.md](README-MY.md))*

An automated **3-phase multi-deck skill** for **Google Antigravity** (and any AI agent supporting `SKILL.md`) that converts source documents (**`.pdf`, `.docx`, `.txt`, `.md`**) into executive **16:9 PowerPoint (`.pptx`)** presentations.

Instead of rewriting boilerplate Python from scratch on every run, this skill bundles **data-driven Python scripts** (`extract.py`, `charts.py`, `build_pptx.py`, `visual_qa.py`) while the AI generates fresh, custom `chart_plan` and `slide_plan` JSONs for every deck.

---

## Key Features

- **Data-Driven Engine (`scripts/`):** Fast, deterministic extraction, 300 DPI transparent-background `matplotlib` charts, and `python-pptx` assembly.
- **10 Consulting & Executive Layouts:** `title_slide`, `section_header`, `key_metrics_cards`, `two_column_chart`, `two_column_image`, `comparison_table`, `executive_summary`, `timeline_roadmap`, `process_flow`, and `matrix_2x2`.
- **4 Built-in Executive Themes + Corporate Brand Kit (`brand.json`):** Customize colors, fonts, company logo, or pass a reference `.pptx` template.
- **Speaker-Notes-Only Citations:** Source references (`[Source: report.pdf, p. 12]`) are placed cleanly in Speaker Notes without cluttering slide visuals.
- **Multi-Deck Isolation & Anti-Overwrite:** Keep multiple decks in separate folders (`slide1/`, `slide2/`); every run outputs a timestamped `Final_Presentation_{YYYYMMDD_HHMMSS}.pptx`.
- **Visual QA Verification:** Checks PPTX XML integrity, slide count, embedded images, and optionally renders PNG previews via LibreOffice + PyMuPDF.

---

## Quick Start (3 Steps)

1. **Open this project** (or install `.agents/skills/pdf-to-pptx/` globally).
2. **Add your source documents** (`.pdf`, `.docx`, `.txt`, `.md`) into `<deck_name>/source/`.
3. Type **`run pdf to pptx`** in chat and answer the 7 short intake questions.

---

## Project Structure

```text
pdf2pptx-presentation/
├── .agents/skills/pdf-to-pptx/
│   ├── SKILL.md                    # Active skill definition
│   ├── scripts/
│   │   ├── extract.py              # Phase 1: PDF/DOCX/TXT/MD + tables + images
│   │   ├── charts.py               # Phase 2: 300 DPI transparent charts
│   │   ├── build_pptx.py           # Phase 3: 16:9 PPTX builder (10 layouts)
│   │   └── visual_qa.py            # QA: XML check + optional PNG render
│   └── references/
│       ├── themes.json             # 4 built-in executive color palettes
│       └── slide_plan_schema.md    # JSON schema reference
├── examples/                       # Sample source + generated demo & previews
├── SKILL.md                        # Root portable copy
├── README.md                       # English documentation
└── README-MY.md                    # Bahasa Melayu documentation
```

---

## Corporate Brand Kit (`brand.json`, Optional)

Drop a `brand.json` file inside `<deck_name>/brand.json` to override any built-in theme:

```json
{
  "company_name": "Acme Advisory",
  "logo_path": "source/logo.png",
  "primary": "#1E3A8A",
  "accent": "#0D9488",
  "background": "#F8FAFC",
  "dark": "#0F172A",
  "body": "#334155",
  "font_title": "Segoe UI",
  "font_body": "Calibri",
  "chart_palette": ["#1E3A8A", "#0D9488", "#E05A47", "#D97706", "#3B82F6", "#475569"]
}
```

---

## Installation Elsewhere

Copy the entire `.agents/skills/pdf-to-pptx/` folder (including `SKILL.md`, `scripts/`, and `references/`) into:

- **Global Antigravity Skills (Recommended):**
  - Windows (PowerShell):
    ```powershell
    Copy-Item -Recurse .\.agents\skills\pdf-to-pptx "$env:USERPROFILE\.gemini\config\skills\pdf-to-pptx" -Force
    ```
  - macOS / Linux:
    ```bash
    mkdir -p ~/.gemini/config/skills && cp -r .agents/skills/pdf-to-pptx ~/.gemini/config/skills/
    ```

---

## Prerequisites

- **Python 3.10+**
- Required Python packages:
  ```bash
  pip install pymupdf pdfplumber pandas matplotlib seaborn python-pptx pillow
  ```
- *(Optional)* **LibreOffice (`soffice`)** on `PATH` for automated PNG slide preview rendering during Visual QA.
