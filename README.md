# PDF-to-PPTX Automated Skill (`SKILL.md`)

Fail **`SKILL.md`** ini ialah satu fail lengkap (*single-file portable skill*) untuk **Google Antigravity** (atau mana-mana ejen AI yang menyokong `SKILL.md` / `GEMINI.md` / `AGENTS.md`). Ia mengautomasikan keseluruhan proses menukar satu atau lebih dokumen sumber (**`.pdf`, `.docx`, `.txt`, `.md`**) kepada pembentangan PowerPoint (`.pptx`) bertahap eksekutif lengkap dengan carta graf **300 DPI latar telus**, **4 tema warna eksekutif**, dan **arahan tambahan (*custom instructions*)**.

---

## Kandungan Dalam Folder Ini

- **`SKILL.md`** — Fail skill utama (1 fail sahaja, boleh disalin terus ke projek atau komputer lain).
- **`README.md`** — Panduan penggunaan dan pemasangan di tempat lain.

---

## Struktur Multi-Deck (Banyak Pembentangan Dalam Satu Projek)

Anda boleh menyimpan dan menjana **banyak pembentangan berasingan** dalam satu projek yang sama dengan meletakkan setiap pembentangan di dalam foldernya sendiri:

```text
SLIDE01/
├── .agents/skills/pdf-to-pptx/SKILL.md   # Skill utama (multi-deck)
├── slide1/                               # Deck 1
│   ├── source/                           # Fail sumber (.pdf, .docx, .txt, .md) & rujukan (.pptx)
│   ├── workspace/                        # Teks ekstrak, tables.json, slide_plan.json
│   ├── assets/                           # Imej ekstrak & carta 300 DPI
│   ├── output/                           # Final_Presentation_<YYYYMMDD_HHMMSS>.pptx
│   ├── extract_pdf.py
│   ├── generate_charts.py
│   └── build_presentation.py
├── slide2/                               # Deck 2
│   └── source/ ...
```

---

## Cara Guna (Aliran Kerja)

1. Letakkan fail sumber anda di dalam `<nama_deck>/source/` (cth: `slide1/source/` atau `slide2/source/`).
2. Taip **arahan pendek** sahaja dalam chat:
   - **Terus nyatakan nama folder deck:** *"Run pdf-to-pptx untuk slide1"* atau *"Buat slide dari folder slide2"*.
   - **Tanpa nyatakan nama folder:** *"Run pdf to pptx"* — Ejen akan imbas folder deck yang ada dan **bertanya soalan pertama: pilih folder deck mana yang hendak dijana** (atau cipta folder baharu).
3. **Fasa 0 (7 Soalan Wajib dalam `<nama_deck>/`):** Ejen akan menyenaraikan fail dalam `<nama_deck>/source/` dan bertanya **7 soalan ringkas**:
   1. **Source data:** Fail sumber utama mana yang hendak digunakan? *(Menyokong `.pdf`, `.docx`, `.txt`, `.md` — boleh lebih daripada 1 fail)*
   2. **Sample rujukan:** Ada fail `.pptx` contoh untuk dirujuk gaya/strukturnya? *(boleh lebih daripada 1 fail, atau `tiada`)*
   3. **Target pages:** Berapa jumlah halaman slide sasaran? *(cth: `15–22 slide`)*
   4. **Maksimum poin per topik:** Berapa maksimum bullet per topik/slide? *(cth: `4 bullet`, huraian penuh masuk ke Speaker Notes)*
   5. **Bahasa output:** Mahu dalam Bahasa Melayu, English, atau Dwi-bahasa?
   6. **Tema Warna (*Color Theme*):** Pilih daripada 4 tema eksekutif siap pakai (**MOH Navy & Teal**, **Corporate Slate & Blue**, **Medical Emerald & Gold**, atau **Executive Burgundy & Amber**) — diselaraskan pada slide dan carta 300 DPI.
   7. **Arahan Tambahan (*Custom Instructions*):** Ada arahan khas? *(Pilihan — cth: fokus pada bab tertentu, nada mesyuarat pengurusan, nama jabatan/pembentang di muka depan, atau `tiada`)*.
4. **Automasi Fasa 1 → Fasa 3 (Terasing dalam `<nama_deck>/`):** Sebaik sahaja anda jawab 7 soalan itu, ejen akan menjalankan semua langkah berikut secara automatik di dalam folder `<nama_deck>/`:
   - **Fasa 1 (`<nama_deck>/extract_pdf.py`):** Mengekstrak teks daripada PDF/DOCX/TXT (`workspace/raw_text.md`), jadual (`workspace/tables.json`), dan imej/rajah (`assets/pdf_img_*.png`).
   - **Fasa 2 (`<nama_deck>/workspace/slide_plan_*.json` & `generate_charts.py`):** Menyusun jalan cerita slide mengikut arahan tambahan & had poin, serta menjana carta graf 300 DPI berlatar telus (`assets/chart_slide_*.png`) mengikut palet tema warna pilihan.
   - **Fasa 3 (`<nama_deck>/build_presentation.py`):** Membina fail akhir **`<nama_deck>/output/Final_Presentation_<YYYYMMDD_HHMMSS>.pptx`** (16:9, susun atur dua kolum, kad KPI, jadual bergaya, nisbah aspek imej terjaga, dan *Speaker Notes*). Setiap larian semula akan menghasilkan fail bertanda masa baharu di dalam folder `output/` supaya versi terdahulu tidak terpadam.

---

## Cara Pasang & Run di Tempat Lain

Anda cuma perlukan fail **`SKILL.md`** ini sahaja. Pilih salah satu daripada 3 cara di bawah:

### Cara 1: Pasang Secara Global (Disyorkan — Berfungsi di Semua Projek)
Salin fail `SKILL.md` ke dalam folder skill global Antigravity di komputer anda:

- **Windows (PowerShell):**
  ```powershell
  mkdir "$env:USERPROFILE\.gemini\config\skills\pdf-to-pptx" -Force
  Copy-Item .\SKILL.md "$env:USERPROFILE\.gemini\config\skills\pdf-to-pptx\SKILL.md"
  ```
- **macOS / Linux:**
  ```bash
  mkdir -p ~/.gemini/config/skills/pdf-to-pptx
  cp SKILL.md ~/.gemini/config/skills/pdf-to-pptx/SKILL.md
  ```
Selepas disalin, anda boleh buka **mana-mana projek** dan terus guna arahan pendek.

### Cara 2: Pasang Dalam Projek Tertentu Sahaja (Boleh Kongsi via Git)
Salin fail `SKILL.md` ke dalam folder `.agents/skills/pdf-to-pptx/` di root projek baharu anda:

- **Windows (PowerShell):**
  ```powershell
  mkdir ".\.agents\skills\pdf-to-pptx" -Force
  Copy-Item .\SKILL.md ".\.agents\skills\pdf-to-pptx\SKILL.md"
  ```
- **macOS / Linux:**
  ```bash
  mkdir -p .agents/skills/pdf-to-pptx
  cp SKILL.md .agents/skills/pdf-to-pptx/SKILL.md
  ```

### Cara 3: Guna Terus Sebagai Rules (`GEMINI.md` / `AGENTS.md`) atau Tampal di Chat
Jika anda bekerja di persekitaran tanpa folder `skills/`, anda boleh:
- Menyalin keseluruhan isi `SKILL.md` (buang bahagian `---` YAML di atas jika perlu) ke dalam fail `GEMINI.md` atau `AGENTS.md` di root projek; **atau**
- Tampal terus isi `SKILL.md` ke dalam chat sebagai arahan permulaan.

---

## Keperluan Sistem (Prerequisites)

- **Python 3.10+**
- Pakej Python berikut (ejen akan semak & pasang secara automatik jika belum ada):
  ```bash
  pip install pymupdf pdfplumber pandas matplotlib seaborn python-pptx pillow
  ```
