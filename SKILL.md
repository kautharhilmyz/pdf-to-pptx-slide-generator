---
name: pdf-to-pptx
description: >-
  Automated 3-phase multi-deck pipeline to convert one or more source documents
  (PDF, DOCX, TXT, MD) inside a chosen deck folder (<deck_dir>/source/) into
  an executive PowerPoint (.pptx) presentation with extracted tables/images,
  300 DPI transparent-background matplotlib charts, 4 executive color themes,
  and optional custom instructions. Use whenever the user gives a short request
  to create slides or a presentation from documents (e.g., "buat slide dari
  pdf", "run pdf-to-pptx untuk slide1", "tukar pdf ke slide"). Always resolves
  the target deck folder first (asking if unspecified), asks 7 intake questions,
  then runs Phases 1–3 automatically inside <deck_dir>/.
---

# PDF-to-PPTX Automated Presentation Pipeline (Multi-Deck)

You are an expert **Python Data Engineer** and **Executive Presentation Architect**.
When this skill is triggered by a short user request, **DO NOT start extracting or generating immediately**. First, complete **Phase 0 (Deck Selection & Pre-Flight Intake)** to resolve `<deck_dir>` and gather the user's exact parameters, and once answered, **automatically execute Phases 1, 2, and 3 end-to-end inside `<deck_dir>/`** without pausing for intermediate confirmations unless a fatal error occurs.

---

## Phase 0: Deck Selection & Mandatory Pre-Flight Intake

### Step 0A: Resolve Target Deck Folder (`<deck_dir>`)
1. **If the user explicitly named a deck folder** in their prompt (e.g., *"run pdf-to-pptx untuk slide1"* or *"buat slide dari folder deck_kardio"*):
   - Use `<deck_dir>` directly (create `<deck_dir>/source/` if needed, or read existing files in `<deck_dir>/source/`).
2. **If the user did NOT specify a deck folder** (e.g., *"run pdf to pptx"*):
   - Scan the workspace root for candidate deck folders (exclude `.agents`, `.vscode`, `.git`, `__pycache__`, `node_modules`).
   - Ask **Question 0 (First Question)** using `ask_question`: *"Which deck folder do you want to generate or update?"* — listing discovered deck folders (plus `Create new deck folder` via custom input, or `.` if using a legacy single-deck root `source/`).
   - Wait for the user's deck choice (or combine Question 0 cleanly if only creating a new deck) before proceeding.

### Step 0B: Cache / Existing Base Data Check (`<deck_dir>/workspace/`)
Before or alongside the intake questions, check if `<deck_dir>/workspace/` already contains previously extracted base data (`raw_text.md`, `tables.json`, `image_manifest.json`):
- **If existing cache/base data is found**, ask the user using `ask_question`:
  *"Terdapat cache / data asas sebelum ini dalam `<deck_dir>/workspace/`. Anda mahu gunakan data sedia ada tersebut atau clear dan jana semula base data yang baru?"*
  - **Use existing cached base data (Recommended if source files unchanged)**: Skip Phase 1 extraction (`extract_pdf.py`) and proceed directly to Phase 2–3 with a new timestamped `slide_plan`, charts, and `.pptx`.
  - **Clear & regenerate fresh base data**: Remove old files in `<deck_dir>/workspace/` and old `pdf_img_*` / `chart_slide_*` in `<deck_dir>/assets/` (preserving `<deck_dir>/output/` history), then re-run Phase 1–3 from `<deck_dir>/source/`.

### Step 0C: 7 Intake Questions (Scoped to `<deck_dir>/`)
Scan `<deck_dir>/source/` to pre-populate file options, then ask the user the following **7 questions** clearly (in the user's language, using `ask_question` for structured choices where helpful, while allowing custom inputs):

1. **Source Data (`source_files`)**: Fail sumber utama yang mana dalam `<deck_dir>/source/`? *(Menyokong multiple files dalam format **`.pdf`**, **`.docx`**, **`.txt`**, atau **`.md`**).*
2. **Sample / Reference Files (`sample_files`)**: Fail contoh atau rujukan yang mana dalam `<deck_dir>/source/`? *(Menyokong multiple `.pptx` files sebagai inspirasi gaya/struktur, atau `tiada` untuk guna tema eksekutif moden secara automatik).*
3. **Target Slide Count (`target_slides`)**: Sasaran jumlah halaman (*target pages*) untuk pembentangan ini? *(cth: `15–22 slide`, `12 slide`, dll.)*
4. **Max Points per Topic (`max_points_per_topic`)**: Maksimum poin/bullet per topik atau per slide? *(Cadangan lalai: maks 4 bullet per slide, 12–15 patah perkataan per bullet; selebihnya masuk ke Speaker Notes atau sub-slide).*
5. **Output Language (`output_language`)**: Bahasa output untuk tajuk, isi slide, label carta, dan *speaker notes*? *(cth: Bahasa Melayu profesional, English, atau Dwi-bahasa).*
6. **Color Theme (`color_theme`)**: Pilih tema warna eksekutif (diselaraskan pada slide & carta 300 DPI):
   - **Classic Navy & Teal** *(Lalai — Background `#F8FAFC`, Dark `#0F172A`, Primary `#1E3A8A`, Accent `#0D9488`, Coral `#E05A47`, Amber `#D97706`)*
   - **Corporate Slate & Blue** *(Background `#F8FAFC`, Dark `#1E293B`, Primary `#0F4C81`, Accent `#0284C7`, Coral `#DC2626`, Amber `#EA580C`)*
   - **Emerald & Gold** *(Background `#F7FEFA`, Dark `#064E3B`, Primary `#047857`, Accent `#D97706`, Coral `#E11D48`, Blue `#0284C7`)*
   - **Executive Burgundy & Amber** *(Background `#FDF8F6`, Dark `#450A0A`, Primary `#7F1D1D`, Accent `#B45309`, Teal `#0F766E`, Blue `#1D4ED8`)*
   *(Jika `sample_files` diberikan, selaraskan jika diminta atau gunakan tema pilihan ini).*
7. **Custom Instructions (`custom_instructions`)**: Ada arahan tambahan khas? *(Pilihan — cth: fokus pada bab tertentu sahaja, nada mesyuarat pengurusan tertinggi, nama jabatan/pembentang di muka depan, atau `tiada`).*

> **Rule:** Wait for the user's answers to these questions. Once answered, proceed automatically through **Environment Setup → Phase 1 → Phase 2 → Phase 3 → Final Verification** strictly inside `<deck_dir>/`.

---

## Environment & Directory Setup

1. Verify Python packages: `pymupdf` (`fitz`), `pdfplumber`, `pandas`, `matplotlib`, `seaborn`, `python-pptx`, and `Pillow` (`PIL`). Install any missing packages via `pip`. (`.docx` extraction uses built-in `zipfile` + `xml.etree.ElementTree`, so no extra package is required.)
2. Create working directories **inside `<deck_dir>/`**:
   - `<deck_dir>/source/` — input source documents and reference `.pptx` files.
   - `<deck_dir>/workspace/` — intermediate extracted text, tables JSON, image manifest, and slide blueprint.
   - `<deck_dir>/assets/` — extracted PDF diagrams/images and generated 300 DPI PNG charts.
   - `<deck_dir>/output/` — final timestamped `.pptx` presentations (`<deck_dir>/output/Final_Presentation_{YYYYMMDD_HHMMSS}.pptx`).

---

## Phase 1: Multi-Format Source Extraction (`<deck_dir>/extract_pdf.py`)

Write and execute `<deck_dir>/extract_pdf.py` (using paths relative to `<deck_dir>/`) to process **all** `source_files` (`.pdf`, `.docx`, `.txt`, `.md`) and inspect `sample_files` (`.pptx`):

1. **Text & Structure Extraction (`<deck_dir>/workspace/raw_text.md`):**
   - **`.pdf`**: Use `pymupdf` (`import pymupdf`) page-by-page, tagged `[FILE: <filename> | PAGE X]`.
   - **`.docx`**: Use Python built-in `zipfile` + `xml.etree.ElementTree` to read `word/document.xml` paragraphs/tables, tagged `[FILE: <filename> | DOCX]`.
   - **`.txt` / `.md`**: Read UTF-8 text directly, tagged `[FILE: <filename> | TEXT]`.
2. **Table Extraction (`<deck_dir>/workspace/tables.json`):**
   - For `.pdf`, use `pdfplumber` to detect and extract numerical/structured tables (`source_file`, `page`, `headers`, `rows`).
3. **Image Extraction (`<deck_dir>/workspace/image_manifest.json`):**
   - For `.pdf`, extract meaningful embedded images/diagrams (`>= 200x200` pixels, excluding extreme aspect-ratio lines and repeated full-page background artwork appearing on >40% of pages). Convert non-RGB/RGBA modes (such as `CMYK`) to `RGB`/`RGBA` before saving.
   - Save valid images as `<deck_dir>/assets/pdf_img_f{file_idx}_p{page}_{img_idx}.png`.
4. **Sample PPTX Inspection (`<deck_dir>/workspace/sample_style.json`, if `sample_files` provided):**
   - Inspect each sample `.pptx` using `python-pptx` and save layout dimensions, color palette, font families, and structural notes.
5. **Phase 1 Self-Check:**
   - Confirm `<deck_dir>/workspace/raw_text.md`, `<deck_dir>/workspace/tables.json`, and `<deck_dir>/workspace/image_manifest.json` exist, then immediately proceed to Phase 2.

---

## Phase 2: Content Synthesis & 300 DPI Chart Generation

Based on `<deck_dir>/workspace/raw_text.md`, `tables.json`, `image_manifest.json`, `sample_style.json`, and the user's `custom_instructions`:

1. **Slide Blueprint (`<deck_dir>/workspace/slide_plan_{YYYYMMDD_HHMMSS}.json`):**
   - Synthesize all source documents into `target_slides` logical, executive-level slides, strictly honoring `custom_instructions` (e.g., chapter focus, audience tone, department/presenter attribution).
   - **Language Rule:** Write **all** slide titles, bullet points, chart titles/labels, table headers, and `speaker_notes` in the user's chosen `output_language`.
   - **Strict Anti-Wall-of-Text & Topic Cap Rule:**
     - Respect `max_points_per_topic` (default max 4 bullets per slide, max 12–15 words per bullet, with **bold lead-ins**).
     - Put detailed explanations, clinical/technical nuances, and citations into `speaker_notes`.
   - **Visual Diversity Rule:**
     - Assign a layout type to every slide: `title_slide`, `section_header`, `two_column_chart`, `two_column_image`, `key_metrics_cards`, or `comparison_table`.
     - Map either an extracted PDF image (`<deck_dir>/assets/pdf_img_...`) OR a newly generated chart (`<deck_dir>/assets/chart_slide_...`) to **at least 70% of content slides**.

2. **Automated Chart Generation (`<deck_dir>/generate_charts.py`):**
   - Identify the **top 5–8 most impactful numerical datasets** from `<deck_dir>/workspace/tables.json` and `raw_text.md`.
   - Generate clean, modern charts (`matplotlib` / `seaborn`: Vertical/Horizontal Bar, Grouped Bar, Line, or Donut — **avoid cluttered 3D pie charts**).
   - **Chart Style Requirements:**
     - `dpi=300` and `transparent=True`
     - Clean sans-serif typography (`Calibri` / `DejaVu Sans` / `Arial`), readable font sizes (11–14pt)
     - Direct data labels on bars/points/slices
     - Remove top and right spines (`ax.spines['top'].set_visible(False)`, `right`)
     - **Theme-Aligned Palette**: Use the exact color palette of the selected `color_theme`:
       - *Classic Navy & Teal*: `["#1E3A8A", "#0D9488", "#E05A47", "#D97706", "#3B82F6", "#475569"]`
       - *Corporate Slate & Blue*: `["#0F4C81", "#0284C7", "#EA580C", "#16A34A", "#6366F1", "#475569"]`
       - *Emerald & Gold*: `["#047857", "#D97706", "#0284C7", "#E11D48", "#0D9488", "#374151"]`
       - *Executive Burgundy & Amber*: `["#7F1D1D", "#B45309", "#0F766E", "#1D4ED8", "#BE123C", "#44403C"]`
   - Save charts as `<deck_dir>/assets/chart_slide_{slide_number}_{YYYYMMDD_HHMMSS}.png` and verify every linked image file exists before Phase 3.

---

## Phase 3: Automated PowerPoint Assembly (`<deck_dir>/build_presentation.py`)

Write and execute `<deck_dir>/build_presentation.py` using `python-pptx` to compile `<deck_dir>/workspace/slide_plan_{YYYYMMDD_HHMMSS}.json` and all visual assets in `<deck_dir>/assets/` into `<deck_dir>/output/Final_Presentation_{YYYYMMDD_HHMMSS}.pptx` (16:9 widescreen: `13.333 x 7.5 inches`).

> **Versioning / Anti-Overwrite Rule:** Every run or improvement (*tambah baik*) MUST append a timestamp suffix (`_YYYYMMDD_HHMMSS`) to the generated `.pptx`, `slide_plan`, and `chart` files. Never overwrite a previous `.pptx` with the same filename.

### Design & Layout Engineering Rules
1. **Theme Application:** Apply the chosen `color_theme` (Background, Dark Title, Primary, Accent, Coral/Amber highlights, Slate body text) consistently across title slides, section dividers, KPI cards, table headers, and accents. Harmonize typography (`Calibri` / `Segoe UI`) with `<deck_dir>/workspace/sample_style.json` if provided.
2. **Dynamic Aspect-Ratio Image/Chart Placement:** Always read exact pixel dimensions with `PIL.Image.open(path).size` and scale proportionally inside the target bounding box — **never stretch or distort aspect ratios**.
3. **Layout Implementation:**
   - `title_slide`, `section_header`, `two_column_chart`, `two_column_image`, `key_metrics_cards`, `comparison_table`.
4. **Typography & Polish:** Generous margins (min `0.6 in`), `1.2x` line spacing, `word_wrap = True`, and inject `speaker_notes` into `slide.notes_slide.notes_text_frame.text`.

---

## Final Verification & Delivery

After running `<deck_dir>/build_presentation.py`:
1. Re-open `<deck_dir>/output/Final_Presentation_{YYYYMMDD_HHMMSS}.pptx` with `python-pptx` to confirm zero XML corruption, exact slide count, all images embedded properly, and speaker notes populated.
2. Report to the user:
   - Exact path of `<deck_dir>/output/Final_Presentation_{YYYYMMDD_HHMMSS}.pptx` and generated charts in `<deck_dir>/assets/`.
   - Concise Markdown summary table (`Slide #`, `Title`, `Layout Type`, `Visual Asset`).
   - Any slides where visuals were omitted or notable assumptions made.
