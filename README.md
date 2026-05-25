# 📋 Exam Outcome Mapper

> **AI-powered exam question ↔ learning outcome alignment tool for accreditation reporting.**  
> Built for Kapadokya University 🏫 — generates ready-to-submit Excel reports 📊 and rubric DOCX files 🗂️ in seconds ⌛️.

🔗 **Live App:** [tekisjet-exammapper.streamlit.app](https://tekisjet-exammapper.streamlit.app)

---

## What It Does

Upload an exam file, define your course learning outcomes, click one button — the AI maps every question to its outcome, classifies Bloom's Taxonomy level and difficulty, then exports a formatted Excel report and a rubric form. The whole process takes under two minutes.

---

## Features

| Feature | Details |
|---|---|
| **AI Matching** | DeepSeek AI maps each question to the most relevant learning outcome |
| **Multi-LO matching** | Questions spanning multiple outcomes get percentage splits (always totalling 100%) |
| **Bloom's Taxonomy** | Auto-classified into all 6 levels (Remember → Create) |
| **Difficulty rating** | Each question rated Easy / Medium / Hard |
| **Bilingual UI** | Full Turkish 🇹🇷 / English 🇬🇧 toggle — all output labels switch with it |
| **Excel export** | 3-sheet workbook: Question Map, LO Summary, Proliz Format |
| **Rubric DOCX** | Formatted rubric form ready for submission |
| **Flexible scoring** | Equal distribution or custom per-question scoring |
| **Answer key** | Optional answer key embedded in the report |
| **Dark / Light mode** | Theme toggle in the sidebar |

---

## Supported File Formats

| Input | Formats |
|---|---|
| Exam file | `.txt`, `.docx`, `.doc`, `.pdf` |
| Learning outcomes | Bilgi Paketi `.pdf` (auto-extract) or manual entry |

---

## Workflow Overview

```
1. Upload exam file  →  questions extracted automatically
2. Load learning outcomes  →  from Bilgi Paketi PDF or manual entry
3. Set scoring  →  equal or custom per question
4. (Optional) Enter answer key + rubric metadata
5. Click "Map with AI"  →  AI runs matching in batches of 20
6. Download Excel report + Rubric DOCX
```

---

## Excel Output — 3 Sheets

### Sheet 1 — Question Map
| Column | Description |
|---|---|
| Q No | Question number |
| Score | Points assigned |
| Question | Full question text |
| LO Rank / ÖÇ Sıra | Matched learning outcome index |
| Weight / Etki Oran | Match percentage (sums to 100) |
| Bloom's Taxonomy | Cognitive level classification |
| Difficulty Level / Zorluk Seviyesi | Easy / Medium / Hard |
| Answer Key / Cevap Anahtarı | Correct answer (if provided) |

### Sheet 2 — LO Summary
Per-outcome summary: how many questions map to each outcome and which question numbers.

### Sheet 3 — Proliz Format
Structured format compatible with university accreditation Proliz system.

---

## Rubric DOCX Output

A print-ready Word document containing:
- Instructor, department, exam type, date header
- Course learning outcomes list
- Scored question table with Bloom level and difficulty
- Bilingual headers matching the selected UI language

---

## Tech Stack

| Component | Technology |
|---|---|
| Frontend | [Streamlit](https://streamlit.io) |
| AI Engine | [DeepSeek Chat API](https://platform.deepseek.com) |
| Excel generation | `openpyxl` |
| DOCX generation | `python-docx` |
| PDF parsing | `pypdf` |
| Hosting | Streamlit Community Cloud |

---

## Setup & Local Development

### 1. Clone the repo
```bash
git clone https://github.com/your-username/exam-outcome-mapper.git
cd exam-outcome-mapper
```

### 2. Install dependencies
```bash
pip install streamlit requests openpyxl python-docx pypdf
```

### 3. Configure secrets
Create `.streamlit/secrets.toml`:
```toml
APP_USERNAME = "your_username"
APP_PASSWORD = "your_password"
DEEPSEEK_API_KEY = "sk-..."
```

### 4. Run
```bash
streamlit run app.py
```

---

## Deployment (Streamlit Community Cloud)

1. Push the repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select your repo and `app-4.py` as the main file
4. Under **Advanced settings → Secrets**, add the three keys above
5. Set a custom subdomain under **Settings → General → Custom subdomain**

---

## Question Extraction Rules

The app auto-detects questions using three strategies (tried in order):

1. Lines matching `1. ... ?` or `1) ... ?` patterns
2. Turkish `Soru 1. ...` format (strips answer choices and answer key)
3. Any line ending with `?` (fallback)

For best results, ensure every question ends with a `?` character.

---

## Learning Outcome Format (Manual Entry)

One outcome per line, colon-separated:
```
LO-1: Explains the fundamental concepts of machine learning
LO-2: Applies supervised learning algorithms to real-world datasets
LO-3: Evaluates model performance using appropriate metrics
```

---

## Language Behavior

Switching between TR / EN changes:
- All UI labels and messages
- Excel column headers (`ÖÇ Sıra` ↔ `LO Rank`, `Etki Oran` ↔ `Weight`)
- LO prefix in outputs (`ÖÇ-4(%100)` ↔ `LO-4(%100)`)
- Bloom taxonomy labels (`Hatırlama` ↔ `Remember`)
- Difficulty labels (`Kolay` ↔ `Easy`)
- Rubric DOCX headers and section titles

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

<p align="center">
  Built with ❤️ at Kapadokya University &nbsp;·&nbsp; Powered by DeepSeek AI<br>
  <strong>by Lecturer Hatice Tekiş 🚀</strong>
</p>
