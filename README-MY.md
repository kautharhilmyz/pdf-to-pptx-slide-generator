# AI Slide Generator (`pdf-to-pptx` Skill)

*(English: see [README.md](README.md))*

Skill automatik **3-fasa multi-deck** untuk **Google Antigravity** (atau mana-mana ejen AI yang menyokong `SKILL.md`) bagi menukar dokumen sumber (**`.pdf`, `.docx`, `.txt`, `.md`**) kepada pembentangan **PowerPoint 16:9 (`.pptx`)** bertahap eksekutif.

Berbanding menulis kod Python dari kosong setiap kali *run*, skill ini membekalkan **enjin skrip berpacukan data** (`extract.py`, `charts.py`, `build_pptx.py`, `visual_qa.py`) sementara ejen AI menjana pelan kandungan (`chart_plan` & `slide_plan` JSON) yang unik untuk setiap pembentangan.

---

## Ciri-Ciri Utama

- **Enjin Berpacukan Data (`scripts/`):** Pengekstrakan pantas, carta `matplotlib` **300 DPI latar telus**, dan pembinaan `.pptx` yang stabil serta jimat token.
- **10 Susun Atur Konsultan & Eksekutif:** `title_slide`, `section_header`, `key_metrics_cards`, `two_column_chart`, `two_column_image`, `comparison_table`, `executive_summary`, `timeline_roadmap`, `process_flow`, dan `matrix_2x2`.
- **4 Tema Warna Eksekutif + Brand Kit Korporat (`brand.json`):** Boleh guna warna syarikat, fon, logo PNG, atau templat `.pptx` rujukan.
- **Sitasi dalam Speaker Notes Sahaja:** Rujukan sumber (`[Source: report.pdf, p. 12]`) dimasukkan hanya dalam *Speaker Notes* supaya slide kekal kemas.
- **Struktur Multi-Deck & Anti-Overwrite:** Asingkan setiap pembentangan dalam folder sendiri (`slide1/`, `slide2/`); setiap larian menghasilkan fail bertanda masa `Final_Presentation_{YYYYMMDD_HHMMSS}.pptx`.
- **Visual QA:** Semak integriti XML `python-pptx`, jumlah slide, imej terbenam, dan (jika LibreOffice dipasang) *render* gambar PNG untuk semakan visual.

---

## Quick Start (3 Langkah)

1. **Buka projek ini** (atau pasang folder `.agents/skills/pdf-to-pptx/` secara global).
2. **Letakkan dokumen sumber** (`.pdf`, `.docx`, `.txt`, `.md`) ke dalam `<nama_deck>/source/`.
3. Taip **`run pdf to pptx`** dalam chat dan jawab 7 soalan ringkas yang dipaparkan.

---

## Struktur Projek

```text
pdf2pptx-presentation/
├── .agents/skills/pdf-to-pptx/
│   ├── SKILL.md                    # Definisi skill aktif
│   ├── scripts/
│   │   ├── extract.py              # Fasa 1: PDF/DOCX/TXT/MD + jadual + imej
│   │   ├── charts.py               # Fasa 2: Carta 300 DPI latar telus
│   │   ├── build_pptx.py           # Fasa 3: Pembina PPTX 16:9 (10 layout)
│   │   └── visual_qa.py            # QA: Semakan XML + render PNG pilihan
│   └── references/
│       ├── themes.json             # 4 palet warna eksekutif siap pakai
│       └── slide_plan_schema.md    # Rujukan skema JSON
├── examples/                       # Contoh sumber + hasil demo & preview
├── SKILL.md                        # Salinan portable di root
├── README.md                       # Dokumentasi Bahasa Inggeris
└── README-MY.md                    # Dokumentasi Bahasa Melayu
```

---

## Brand Kit Korporat (`brand.json`, Pilihan)

Letakkan fail `brand.json` di dalam `<nama_deck>/brand.json` untuk menggunakan identiti jenama sendiri:

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

## Cara Pasang di Projek / Komputer Lain

Salin keseluruhan folder `.agents/skills/pdf-to-pptx/` (termasuk `SKILL.md`, `scripts/`, dan `references/`) ke:

- **Global Antigravity Skills (Disyorkan):**
  - Windows (PowerShell):
    ```powershell
    Copy-Item -Recurse .\.agents\skills\pdf-to-pptx "$env:USERPROFILE\.gemini\config\skills\pdf-to-pptx" -Force
    ```
  - macOS / Linux:
    ```bash
    mkdir -p ~/.gemini/config/skills && cp -r .agents/skills/pdf-to-pptx ~/.gemini/config/skills/
    ```

---

## Keperluan Sistem (Prerequisites)

- **Python 3.10+**
- Pakej Python yang diperlukan:
  ```bash
  pip install pymupdf pdfplumber pandas matplotlib seaborn python-pptx pillow
  ```
- *(Pilihan)* **LibreOffice (`soffice`)** dalam `PATH` untuk *render* gambar preview PNG secara automatik semasa Visual QA.
