# 📋 Exam Outcome Mapper — User Guide

**URL:** [tekisjet-exammapper.streamlit.app](https://tekisjet-exammapper.streamlit.app)

---

## Table of Contents

1. [Logging In](#1-logging-in)
2. [Interface Overview](#2-interface-overview)
3. [Step 1 — Upload Your Exam File](#3-step-1--upload-your-exam-file)
4. [Step 2 — Define Learning Outcomes](#4-step-2--define-learning-outcomes)
5. [Step 3 — Set Scoring](#5-step-3--set-scoring)
6. [Answer Key (Optional)](#6-answer-key-optional)
7. [Rubric Form Details](#7-rubric-form-details)
8. [Step 4 — Generate Report](#8-step-4--generate-report)
9. [Understanding Your Outputs](#9-understanding-your-outputs)
10. [Language & Theme](#10-language--theme)
11. [Tips & Common Issues](#11-tips--common-issues)

---

## 1. Logging In

Open the app URL. You will see a login screen.

Enter the **username** and **password** provided by your administrator, then click **Login**.

> If you see ❌ *Invalid username or password*, double-check your credentials. Contact your system admin if the issue persists.

---

## 2. Interface Overview

After logging in you will see:

- **Sidebar (left panel)** — language toggle, theme toggle, Start Over, Logout
- **Main area** — four steps running top to bottom
- **Download buttons** — appear after the AI mapping is complete

### Sidebar Controls

| Control | What it does |
|---|---|
| 🌐 TR / EN | Switch the entire UI and all output labels between Turkish and English |
| 🌙 Dark / ☀️ Light | Toggle dark or light theme |
| 🔄 Start Over | Resets the page without clearing your session (re-upload to refresh) |
| 🚪 Logout | Returns you to the login screen |

---

## 3. Step 1 — Upload Your Exam File

Click the upload area (or drag and drop) to upload your exam.

**Accepted formats:** `.txt`, `.docx`, `.doc`, `.pdf`

The app scans the file and extracts questions automatically. You will see:

- ✅ **X questions found** — extraction succeeded, click the Preview toggle to verify
- ❌ **No questions found** — see tips below

### How Questions Are Detected

The app recognises three common formats:

**Format A — Numbered with punctuation**
```
1. What is the purpose of the Flatten() layer?
2. Which activation function is used for multi-class output?
```

**Format B — Turkish "Soru" prefix**
```
Soru 1. Aşağıdaki Keras kodu hangi tür katman ekler?
Soru 2. model.fit() çağrısında 'validation_split=0.2' ne anlama gelir?
```

**Format C — Any line ending with `?`**
```
What does Conv2D(32, (3,3)) do?
Why is softmax used in the output layer?
```

> **Important:** Every question must end with a `?` mark. Answer choices (A, B, C…) and answer keys are automatically stripped.

---

## 4. Step 2 — Define Learning Outcomes

You have two options — choose the tab that suits you:

### Option A: Extract from PDF (Bilgi Paketi)

1. Click **"Extract from PDF"** tab
2. Upload your course information package (Bilgi Paketi) PDF
3. The AI reads the PDF and pulls out all learning outcomes automatically
4. You will see them listed with their numbers and definitions

> This works best with standard university Bilgi Paketi documents. The AI looks for the "Dersin Öğrenme Çıktıları" section.

### Option B: Enter Manually

1. Click **"Enter Manually"** tab
2. Type or paste your learning outcomes, one per line
3. Use the format: `LO-1: definition text here`

**Example:**
```
LO-1: Explains the fundamental concepts of convolutional neural networks
LO-2: Applies CNN models to image classification tasks
LO-3: Evaluates model performance using loss and accuracy metrics
LO-4: Designs and trains deep learning models using Keras
```

Once outcomes are loaded you will see: ✅ **X outcomes defined**

---

## 5. Step 3 — Set Scoring

### Total Exam Points

Set the total points for the exam (default: 100).

### Scoring Type

**Equal (auto)** — Every question gets the same points.  
`Total ÷ Number of questions`  
Example: 100 points ÷ 50 questions = 2.00 pts each

**Custom per question** — Individual point boxes appear for each question.  
Enter the score for each one. The app shows a running total:
- ⚠️ Total: 98 — means your scores don't add up to the exam total yet
- ✅ Total: 100 — all good

---

## 6. Answer Key (Optional)

Expand the **Answer Key** section and type the correct answers as a continuous string — no spaces or separators.

**Example for 10 questions:**
```
ABCDEABCDE
```

The answer key is embedded in the Excel report in the rightmost column of Sheet 1. Leave it blank if not needed.

---

## 7. Rubric Form Details

Fill in the four fields to personalise the rubric DOCX header:

| Field | Example |
|---|---|
| Instructor | Öğr. Gör. Hatice Tekiş |
| Exam Type | Vize / Final / Quiz |
| Department / Program | Bilgisayar Programcılığı |
| Exam Date | 01/06/2026 |

These appear at the top of the downloaded rubric document. All fields are optional — leave blank if not needed.

---

## 8. Step 4 — Generate Report

Once you have completed Steps 1 and 2, the **Map with AI & Generate Excel** button becomes active.

Click it. The AI processes questions in batches of 20. A progress bar shows which batch is running.

**What the AI decides for each question:**
- Which learning outcome(s) it maps to
- If it spans multiple outcomes, what percentage belongs to each (always totals 100%)
- Which Bloom's Taxonomy level it tests
- Whether it is Easy, Medium, or Hard

When done you will see:
> ✅ **52/52 questions mapped — 3 matched to multiple outcomes**

Two download buttons then appear:

| Button | File | Contents |
|---|---|---|
| ⬇ Download Excel Report | `ExamName-Outcome-Report.xlsx` | 3-sheet workbook |
| 📄 Download Rubric Form | `ExamName-Rubrik.docx` | Print-ready rubric table |

---

## 9. Understanding Your Outputs

### Excel Workbook — 3 Sheets

**Sheet 1: Question Map**

The main mapping table. Each row is one question.

| Column | Description |
|---|---|
| Q No | Question number |
| Score | Points for that question |
| Question | Full question text |
| LO Rank 1 / ÖÇ Sıra 1 | Index number of the primary matched outcome |
| Weight 1 / Etki Oran 1 | Percentage of the match (e.g. 60) |
| LO Rank 2 / ÖÇ Sıra 2 | Second outcome index (if multi-match) |
| Weight 2 / Etki Oran 2 | Second outcome percentage (e.g. 40) |
| Bloom's Taxonomy | Cognitive level |
| Difficulty | Easy / Medium / Hard |
| Answer Key | Correct answer letter (if provided) |

> When two outcomes are matched, the weights always sum to 100. For example: 60 + 40, or 70 + 30.

**Sheet 2: LO Summary**

A quick overview per learning outcome:
- How many questions map to it
- Which question numbers those are

Useful for checking whether any outcome has too few or too many questions assigned.

**Sheet 3: Proliz Format**

The same data structured for the Proliz accreditation system. Same columns as Sheet 1 but with Turkish field names regardless of language setting.

---

### Rubric DOCX

A Word document with two parts:

1. **Header table** — Instructor, exam type, department, date, course name
2. **Learning outcomes list** — all LOs numbered
3. **Scoring rubric table** — one row per question with:
   - Question number and points
   - Correct / wrong answer columns (pre-filled with 4 / 0)
   - Question text
   - Matched learning outcome(s) with percentages
   - Bloom's Taxonomy level
   - Difficulty

---

### Bloom's Taxonomy Levels Explained

| Level | TR | Typical question cues |
|---|---|---|
| Remember | Hatırlama | "What is…", "Which…", "Define…", "List…" |
| Understand | Anlama | "Explain…", "Describe…", "What does X mean…" |
| Apply | Uygulama | "Which code does X", "Calculate…", "Use X to…" |
| Analyze | Analiz | "Compare…", "Why does X cause Y…", "Differentiate…" |
| Evaluate | Değerlendirme | "Which approach is better…", "Justify…", "Assess…" |
| Create | Yaratma | "Design…", "Build…", "Formulate…", "Propose…" |

---

### Difficulty Levels

| Level | TR | Meaning |
|---|---|---|
| Easy | Kolay | Straightforward recall or recognition |
| Medium | Orta | Requires understanding or simple application |
| Hard | Zor | Requires analysis, evaluation, or creation |

---

## 10. Language & Theme

### Switching Language

Use the **TR / EN** toggle in the sidebar. The switch takes effect immediately and affects:

- All UI text and labels
- Excel column headers (`ÖÇ Sıra` ↔ `LO Rank`)
- LO prefix in outputs (`ÖÇ-4(%100)` ↔ `LO-4(%100)`)
- Bloom level labels (`Hatırlama` ↔ `Remember`)
- Difficulty labels (`Kolay` ↔ `Easy`)
- Rubric DOCX headers and titles

> Switch the language **before** running the mapping, so that the downloaded files use your preferred language.

### Dark / Light Mode

Click the **🌙 Dark** or **☀️ Light** button in the sidebar to toggle the theme. This is a visual preference only and does not affect the output files.

---

## 11. Tips & Common Issues

### ❌ No questions found

- Make sure every question ends with `?`
- Try converting your file to `.txt` and re-uploading
- If using a PDF, check that the text is selectable (not a scanned image)
- For DOCX files with tables, the app reads table cells automatically

### ❌ No outcomes found (PDF extraction)

- Make sure the PDF is a standard Bilgi Paketi — the AI looks for the section heading "Dersin Öğrenme Çıktıları"
- If extraction fails, switch to **Enter Manually** tab and paste your outcomes

### ⚠️ Custom scoring total doesn't match

- The sum of all per-question scores must equal the Total Exam Points you set
- The app highlights the discrepancy in real time — adjust individual scores until you see ✅

### AI mapping seems slow

- The AI processes 20 questions per batch — a 60-question exam runs 3 batches
- Each batch typically takes 5–15 seconds
- Do not close or reload the tab while mapping is running

### Download buttons disappeared after I did something

- Click **🔄 Start Over** in the sidebar and re-run the mapping
- Session state is preserved until you refresh the browser or log out

### I want to re-run the mapping with different outcomes

- Update your outcomes in Step 2
- Click **Map with AI** again — the new results will replace the previous ones

---

*For technical issues or access problems, contact your system administrator.*

---

<p align="center">by Lecturer Hatice Tekiş &nbsp;·&nbsp; Kapadokya University</p>
