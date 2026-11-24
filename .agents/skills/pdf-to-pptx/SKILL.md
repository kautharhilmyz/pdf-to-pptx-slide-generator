---
name: pdf-to-pptx
description: >-
  Automated 3-phase multi-deck pipeline to convert source documents (PDF, DOCX,
  TXT, MD) inside a chosen deck folder (<deck_dir>/source/) into an executive
  16:9 PowerPoint (.pptx) presentation using bundled data-driven Python
  scripts, 300 DPI transparent-background charts, 10 consulting layouts, 4
  built-in themes + corporate brand.json support, and Speaker-Notes-only source
  citations. Use whenever the user requests slides or a presentation from
  documents (e.g., "convert pdf to pptx", "run pdf-to-pptx for slide1", "buat
  slide dari pdf"). Always resolves the target deck folder first, asks 7 intake
  questions, then runs Phases 1-3 automatically inside <deck_dir>/.
---

# PDF-to-PPTX Automated Presentation Pipeline (Multi-Deck)

You are an expert **Python Data Engineer** and **Executive Presentation Architect**.
When this skill is triggered, **DO NOT start extracting or generating immediately**. First complete **Phase 0 (Deck Selection & Pre-Flight Intake)** to resolve `<deck_dir>` and gather parameters, then execute **Phases 1–3** end-to-end inside `<deck_dir>/` using the bundled scripts in `scripts/`.

---

## Phase 0: Deck Selection & Mandatory Pre-Flight Intake

### Step 0A: Resolve Target Deck Folder (`<deck_dir>`)
1. **If the user explicitly named a deck folder** (e.g., *"run pdf-to-pptx for slide1"*):
   - Use `<deck_dir>` directly (create `<deck_dir>/source/` if missing).
2. **If the user did NOT specify a deck folder** (e.g., *"run pdf to pptx"*):
   - Scan the workspace root for candidate deck folders (exclude `.agents`, `.vscode`, `.git`, `__pycache__`, `node_modules`, `examples`).
   - Ask **Question 0** using `ask_question`: *"Which deck folder do you want to generate or update?"* — listing discovered folders plus an option to create a new deck folder.

### Step 0B: Cache / Existing Base Data Check (`<deck_dir>/workspace/`)
Check if `<deck_dir>/workspace/` already contains `raw_text.md`, `tables.json`, and `image_manifest.json`:
- **If cached data exists**, ask via `ask_question`:
  *"Existing extracted base data was found in `<deck_dir>/workspace/`. Use cached data or clear and re-extract?"*
  - **Use existing cached base data (Recommended if source files unchanged)**: Skip Phase 1 and proceed directly to Phase 2–3 with new timestamped `chart_plan`, `slide_plan`, charts, and `.pptx`.
  - **Clear & regenerate fresh base data**: Remove old files in `<deck_dir>/workspace/` and old `pdf_img_*` / `chart_slide_*` in `<deck_dir>/assets/` (preserving `<deck_dir>/output/`), then run Phases 1–3.

### Step 0C: 7 Intake Questions (Scoped to `<deck_dir>/`)
Scan `<deck_dir>/source/` to pre-populate file options, then ask the user these **7 questions** (in the user's language, using `ask_question` where helpful):

1. **Source Data (`source_files`)**: Which primary source files in `<deck_dir>/source/`? *(Supports multiple `.pdf`, `.docx`, `.txt`, `.md` files.)*
2. **Sample / Reference Files (`sample_files`)**: Any reference `.pptx` files for style/template inspiration? *(Supports multiple `.pptx` files, or `none`.)*
3. **Target Slide Count (`target_slides`)**: Target number of slides? *(e.g., `12–18 slides`.)*
4. **Max Points per Topic (`max_points_per_topic`)**: Maximum bullets per slide? *(Default: max 4 bullets per slide, 12–15 words per bullet; extra detail goes into Speaker Notes.)*
5. **Output Language (`output_language`)**: Output language for slide titles, bullets, chart labels, and speaker notes? *(e.g., English, Bahasa Melayu, or Bilingual.)*
6. **Color Theme / Brand Kit (`color_theme`)**: Choose a theme (or use `<deck_dir>/brand.json` if present):
   - **Classic Navy & Teal** (`classic_navy_teal` — Default)
   - **Corporate Slate & Blue** (`corporate_slate_blue`)
   - **Emerald & Gold** (`emerald_gold`)
   - **Executive Burgundy & Amber** (`executive_burgundy_amber`)
7. **Custom Instructions (`custom_instructions`)**: Any special instructions? *(Optional — e.g., focus on specific chapters, executive tone, presenter/department name, or `none`.)*

> **Rule:** Wait for answers before running Phases 1–3. Once answered, execute Phases 1–3 automatically inside `<deck_dir>/`.

---

## Environment & Directory Setup

1. Verify Python packages: `pymupdf`, `pdfplumber`, `pandas`, `matplotlib`, `seaborn`, `python-pptx`, `Pillow`. Install missing packages via `pip` if needed.
2. Ensure `<deck_dir>/source/`, `<deck_dir>/workspace/`, `<deck_dir>/assets/`, and `<deck_dir>/output/` exist.
3. Locate the skill's `scripts/` directory (`.agents/skills/pdf-to-pptx/scripts/` or global skill folder).

---

## Phase 1: Multi-Format Source Extraction (`scripts/extract.py`)

Run the bundled extraction script (unless using cached data):
```bash
python <skill_dir>/scripts/extract.py --deck-dir "<deck_dir>" --sources <file1> <file2> [--samples <sample1.pptx>]
```
This produces:
- `<deck_dir>/workspace/raw_text.md`
- `<deck_dir>/workspace/tables.json`
- `<deck_dir>/workspace/image_manifest.json`
- `<deck_dir>/workspace/sample_style.json` (if `--samples` provided)

Confirm these files exist, then proceed immediately to Phase 2.

---

## Phase 2: Content Synthesis & 300 DPI Chart Generation (`scripts/charts.py`)

Read `<deck_dir>/workspace/raw_text.md`, `tables.json`, `image_manifest.json`, and `references/slide_plan_schema.md`.
Let `{ts}` be the current timestamp `YYYYMMDD_HHMMSS`.

1. **Generate Chart Plan (`<deck_dir>/workspace/chart_plan_{ts}.json`):**
   - Identify the **top 5–8 most impactful numerical datasets** from `tables.json` and `raw_text.md`.
   - Write `<deck_dir>/workspace/chart_plan_{ts}.json` following `references/slide_plan_schema.md` (`bar`, `barh`, `line`, `donut`).
   - Run:
     ```bash
     python <skill_dir>/scripts/charts.py --deck-dir "<deck_dir>" --plan "workspace/chart_plan_{ts}.json"
     ```
2. **Generate Slide Blueprint (`<deck_dir>/workspace/slide_plan_{ts}.json`):**
   - Synthesize `target_slides` slides in `output_language`, honoring `custom_instructions` and `max_points_per_topic`.
   - **Concise Topic-Focused Titles:** Keep slide titles short and focused (3–7 words) so they frame the topic cleanly without wrapping across multiple lines.
   - **Anti-Wall-of-Text Rule:** Max 4 bullets per slide, 12–15 words per bullet with **bold lead-ins** (`**Lead-in:** detail`). Put deeper explanations in `speaker_notes`.
   - **Speaker-Notes-Only Citations:** Place all source citations (e.g., `[Source: report.pdf, p. 12]`) **strictly inside `speaker_notes`** — never on the visible slide canvas.
   - **Visual Diversity (10 Layouts):** Use `title_slide`, `section_header`, `two_column_chart`, `two_column_image`, `key_metrics_cards`, `comparison_table`, `executive_summary`, `timeline_roadmap`, `process_flow`, and `matrix_2x2`. Attach a chart or extracted image to **at least 70% of content slides** where relevant.

---

## Phase 3: Automated PowerPoint Assembly & QA (`scripts/build_pptx.py`, `scripts/visual_qa.py`)

1. **Build Presentation:**
   ```bash
   python <skill_dir>/scripts/build_pptx.py --deck-dir "<deck_dir>" --plan "workspace/slide_plan_{ts}.json" --output "output/Final_Presentation_{ts}.pptx"
   ```
   > **Anti-Overwrite Rule:** Always use a unique `{ts}` (`YYYYMMDD_HHMMSS`) suffix for `chart_plan`, `slide_plan`, `chart_slide_*`, and `Final_Presentation_*.pptx`.
2. **Verify & Optional Visual QA:**
   ```bash
   python <skill_dir>/scripts/visual_qa.py --pptx "<deck_dir>/output/Final_Presentation_{ts}.pptx" --render-dir "<deck_dir>/workspace/qa_previews"
   ```
   - Confirm zero XML corruption, exact slide count, embedded images, and populated speaker notes.
   - If `visual_qa_rendered` is `true` (LibreOffice installed), inspect the rendered PNGs in `<deck_dir>/workspace/qa_previews/` to verify no text overlap or clipping.

---

## Final Delivery

Report to the user:
- Exact path of `<deck_dir>/output/Final_Presentation_{ts}.pptx` and generated charts in `<deck_dir>/assets/`.
- Concise Markdown summary table (`Slide #`, `Title`, `Layout`, `Visual Asset`).
- Notable assumptions or omitted visuals, if any.
