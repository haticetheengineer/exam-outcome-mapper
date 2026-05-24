import streamlit as st
import json
import re
import io
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Exam Outcome Mapper", page_icon="📋", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.header {
    background: linear-gradient(135deg, #1a3a6b, #2c5aa0);
    color: white; padding: 20px; border-radius: 10px; margin-bottom: 24px; text-align: center;
}
.header h1 { margin:0; font-size:1.4rem; }
.header p  { margin:6px 0 0; font-size:0.8rem; opacity:0.8; }
.login-box {
    background: white; border: 1px solid #e0e0e0;
    border-radius: 12px; padding: 32px; max-width: 400px;
    margin: 60px auto; box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}
.login-box h2 { text-align:center; margin-bottom:24px; color:#1a3a6b; font-size:1.2rem; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# GİRİŞ EKRANI
# ══════════════════════════════════════════════════════════════
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown("""<div class='header'>
        <h1>📋 Exam Outcome Mapper</h1>
        <p>Kapadokya University — Accreditation Report System</p>
    </div>""", unsafe_allow_html=True)

    with st.container():
        st.markdown("<div class='login-box'>", unsafe_allow_html=True)
        st.markdown("### 🔐 Login")
        username = st.text_input("Username", placeholder="Enter username")
        password = st.text_input("Password", type="password", placeholder="Enter password")

        if st.button("Login", type="primary", use_container_width=True):
            valid_user = st.secrets.get("APP_USERNAME", "admin")
            valid_pass = st.secrets.get("APP_PASSWORD", "admin123")
            if username == valid_user and password == valid_pass:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("❌ Invalid username or password")
        st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# ══════════════════════════════════════════════════════════════
# ANA UYGULAMA
# ══════════════════════════════════════════════════════════════
st.markdown("""<div class='header'>
    <h1>📋 Exam Outcome Mapper</h1>
    <p>Kapadokya University — Accreditation Report System</p>
</div>""", unsafe_allow_html=True)

# Logout
with st.sidebar:
    st.markdown("### 📋 Exam Outcome Mapper")
    st.markdown("---")
    st.markdown(f"👤 Logged in as **{st.secrets.get('APP_USERNAME','admin')}**")
    if st.button("🔄 Start Over", use_container_width=True):
        st.rerun()
    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()
    st.markdown("---")
    st.markdown("<small style='color:#aaa'>Kapadokya University<br>🤖 DeepSeek AI</small>", unsafe_allow_html=True)

DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"

def deepseek_chat(messages, max_tokens=4096):
    api_key = st.secrets["DEEPSEEK_API_KEY"]
    resp = requests.post(
        DEEPSEEK_URL,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={"model": "deepseek-chat", "messages": messages, "max_tokens": max_tokens},
        timeout=60
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"].strip()

def extract_questions(text):
    sorular = []

    # Format 1: Numaralı sorular "?" ile biter (normal döküman)
    pattern1 = r'(\d{1,3}[\.\)]\s*)([\s\S]*?\?)'
    for idx, (_, soru) in enumerate(re.findall(pattern1, text), 1):
        soru = re.sub(r'\s+', ' ', soru.strip())
        if len(soru) > 8:
            sorular.append({"no": idx, "text": soru})
    if sorular:
        return sorular

    # Format 2: "Soru X. soru metni" — tablo formatı
    pattern2 = r'Soru\s+(\d+)\.\s+(.*?)(?=Soru\s+\d+\.|$)'
    matches2 = re.findall(pattern2, text, re.DOTALL)
    if matches2:
        for no_str, metin in matches2:
            metin = re.sub(r'\n\s*[A-E]\)\s.*', '', metin)
            metin = re.sub(r'Cevap:.*', '', metin, flags=re.DOTALL)
            metin = re.sub(r'\|', '', metin)
            metin = re.sub(r'\s+', ' ', metin.strip())
            if len(metin) > 8:
                sorular.append({"no": int(no_str), "text": metin})
        if sorular:
            return sorular

    # Format 3: Satır sonunda "?" olan satırlar
    for idx, line in enumerate(text.split('\n'), 1):
        line = line.strip()
        if line.endswith('?') and len(line) > 8:
            sorular.append({"no": idx, "text": line})

    return sorular

def extract_oc_from_pdf(pdf_bytes):
    """PDF'den ÖÇ'leri çıkar — DeepSeek ile"""
    import pypdf
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    text = "\n".join(p.extract_text() or "" for p in reader.pages)

    prompt = f"""This is a Turkish university course information package (bilgi paketi).
Extract all learning outcomes (Öğrenme Çıktıları / Dersin Öğrenme Çıktıları) from the text.
They are usually labeled as Ö01, Ö02... or numbered list under "Dersin Öğrenme Çıktıları".

Return ONLY this JSON format:
{{"outcomes":[{{"no":"LO-1","definition":"..."}},{{"no":"LO-2","definition":"..."}}]}}

TEXT:
{text[:8000]}"""

    raw = deepseek_chat([{"role": "user", "content": prompt}], max_tokens=2048)
    raw = re.sub(r'```json|```', '', raw).strip()
    m = re.search(r'\{[\s\S]*\}', raw)
    if m:
        data = json.loads(m.group())
        return [{"no": o["no"], "tanim": o["definition"]} for o in data.get("outcomes", [])]
    return []

def match_and_build_excel(sorular, ocler, anahtar=""):
    oc_listesi = "\n".join([f"- {o['no']}: {o['tanim']}" for o in ocler])
    eslestirmeler = {}

    n = len(sorular)
    chunk_size = 20
    chunks = (n + chunk_size - 1) // chunk_size
    progress = st.progress(0, text="AI mapping questions...")

    for ci, i in enumerate(range(0, n, chunk_size)):
        chunk = sorular[i:i+chunk_size]
        soru_listesi = "\n".join([f"{s['no']}. {s['text']}" for s in chunk])

        prompt = f"""You are an academic assessment expert. Match each exam question to the most appropriate learning outcome.

LEARNING OUTCOMES:
{oc_listesi}

EXAM QUESTIONS:
{soru_listesi}

Rules:
- Every question MUST be matched to exactly one learning outcome
- Choose the outcome that best aligns with the question topic
- Assign difficulty: Easy, Medium, or Hard

Return ONLY valid JSON:
{{"matches":[{{"q":1,"lo":"LO-1","difficulty":"Medium"}}]}}"""

        try:
            raw = deepseek_chat([{"role": "user", "content": prompt}], max_tokens=2048)
            raw = re.sub(r'```json|```', '', raw).strip()
            m = re.search(r'\{[\s\S]*\}', raw)
            if m:
                data = json.loads(m.group())
                for item in data.get("matches", []):
                    q_no = int(item.get("q", 0))
                    lo   = str(item.get("lo", ""))
                    diff = str(item.get("difficulty", "Medium"))
                    if q_no > 0 and lo:
                        eslestirmeler[q_no] = {"oc_no": lo, "zorluk": diff}
        except Exception as e:
            st.warning(f"Batch {ci+1} error: {e}")

        progress.progress((ci+1)/chunks, text=f"Mapping batch {ci+1}/{chunks}...")

    progress.empty()

    # Excel oluştur
    wb = Workbook()
    BLUE="1A3A6B"; GREEN="1E6B3A"; GREEN_L="E8F5ED"
    RED="C0392B"; GRAY="F5F5F5"; WHITE="FFFFFF"

    def hstyle(cell):
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def add_borders(ws, r1, r2, c1, c2):
        t = Side(style="thin", color="CCCCCC")
        for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
            for cell in row:
                cell.border = Border(left=t, right=t, top=t, bottom=t)

    oc_map = {o["no"]: o["tanim"] for o in ocler}

    # Sayfa 1: Question-LO Mapping
    ws1 = wb.active
    ws1.title = "Question-LO Mapping"
    ws1.row_dimensions[1].height = 32
    for col, h in enumerate(["Q NO","QUESTION","ANSWER KEY","LO NO","LO DEFINITION","DIFFICULTY"], 1):
        hstyle(ws1.cell(row=1, column=col, value=h))

    for ri, s in enumerate(sorular, 2):
        esl  = eslestirmeler.get(s["no"], {})
        lo   = esl.get("oc_no", "")
        diff = esl.get("zorluk", "")
        key  = anahtar[s["no"]-1] if anahtar and s["no"]-1 < len(anahtar) else "-"

        ws1.cell(ri,1,s["no"]).alignment = Alignment(horizontal="center")
        ws1.cell(ri,2,s["text"]).alignment = Alignment(wrap_text=True, vertical="top")
        ws1.cell(ri,3,key).alignment = Alignment(horizontal="center")
        ws1.cell(ri,4,lo).alignment = Alignment(horizontal="center")
        ws1.cell(ri,5,oc_map.get(lo,"")).alignment = Alignment(wrap_text=True)
        ws1.cell(ri,6,diff).alignment = Alignment(horizontal="center")

        if lo:
            for c in range(1,7):
                ws1.cell(ri,c).fill = PatternFill("solid", fgColor="F0FFF4")
        elif ri % 2 == 0:
            for c in range(1,7):
                ws1.cell(ri,c).fill = PatternFill("solid", fgColor=GRAY)

    for col, w in zip("ABCDEF", [8, 60, 12, 10, 46, 12]):
        ws1.column_dimensions[col].width = w
    ws1.freeze_panes = "A2"
    add_borders(ws1, 1, len(sorular)+1, 1, 6)

    # Sayfa 2: LO Summary
    ws2 = wb.create_sheet("LO Summary")
    ws2.row_dimensions[1].height = 32
    for col, h in enumerate(["LO NO","LO DEFINITION","# QUESTIONS","QUESTION NUMBERS"],1):
        hstyle(ws2.cell(row=1, column=col, value=h))

    for ri, oc in enumerate(ocler, 2):
        matched = [s for s in sorular if eslestirmeler.get(s["no"],{}).get("oc_no")==oc["no"]]
        q_nums  = ", ".join(str(s["no"]) for s in matched)

        ws2.cell(ri,1,oc["no"]).alignment = Alignment(horizontal="center")
        ws2.cell(ri,2,oc["tanim"]).alignment = Alignment(wrap_text=True)
        ws2.cell(ri,3,len(matched)).alignment = Alignment(horizontal="center")
        ws2.cell(ri,4,q_nums).alignment = Alignment(wrap_text=True)

        if ri%2==0:
            for c in [1,2,3,4]:
                ws2.cell(ri,c).fill = PatternFill("solid",fgColor=GRAY)

    for col, w in zip("ABCD",[10,52,14,40]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"
    add_borders(ws2, 1, len(ocler)+1, 1, 4)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), len(eslestirmeler)

# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════

# BÖLÜM 1: Sınav Dosyası
st.markdown("### 📂 Step 1 — Upload Exam File")
uploaded_exam = st.file_uploader("TXT, DOCX or PDF", type=["txt","pdf","docx","doc"], key="exam")

sorular = []
exam_filename = "Exam-Outcome-Report"
if uploaded_exam:
    exam_filename = uploaded_exam.name.rsplit(".", 1)[0]  # uzantısız isim
    ext = uploaded_exam.name.split(".")[-1].lower()
    raw = uploaded_exam.read()
    try:
        if ext == "txt":
            text = raw.decode("utf-8", errors="ignore")
            sorular = extract_questions(text)
        elif ext in ["docx","doc"]:
            from docx import Document
            doc = Document(io.BytesIO(raw))
            # Paragrafları oku
            para_text = "\n".join(p.text for p in doc.paragraphs)
            # Tablolardaki metni ayrı topla (tekrar önlemek için set kullan)
            table_text = ""
            seen_cells = set()
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        ct = cell.text.strip()
                        if ct and ct not in seen_cells:
                            seen_cells.add(ct)
                            table_text += ct + "\n"
            # Önce paragraflardan dene, bulamazsa tablodan
            sorular = extract_questions(para_text)
            if not sorular:
                sorular = extract_questions(table_text)
        elif ext == "pdf":
            import pypdf
            text = "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(io.BytesIO(raw)).pages)
            sorular = extract_questions(text)

        if sorular:
            st.success(f"✅ {len(sorular)} questions found")
            with st.expander("Preview questions", expanded=False):
                for s in sorular:
                    st.markdown(f"**{s['no']}.** {s['text']}")
        else:
            st.error("❌ No questions found. Questions must end with '?'")
    except Exception as e:
        st.error(f"❌ {e}")

st.divider()

# BÖLÜM 2: ÖÇ
st.markdown("### 🎯 Step 2 — Learning Outcomes")

tab1, tab2 = st.tabs(["📄 Extract from PDF (Bilgi Paketi)", "✏️ Enter Manually"])

ocler = []
oc_text_default = ""

with tab1:
    st.info("Upload your course information package (Bilgi Paketi PDF) — outcomes will be extracted automatically.")
    uploaded_bp = st.file_uploader("Bilgi Paketi PDF", type=["pdf"], key="bilgi")
    if uploaded_bp:
        with st.spinner("Extracting learning outcomes from PDF..."):
            try:
                extracted = extract_oc_from_pdf(uploaded_bp.read())
                if extracted:
                    st.success(f"✅ {len(extracted)} outcomes extracted!")
                    oc_text_default = "\n".join([f"{o['no']}: {o['tanim']}" for o in extracted])
                    for oc in extracted:
                        st.markdown(f"**{oc['no']}** — {oc['tanim'][:80]}...")
                    ocler = extracted
                else:
                    st.error("❌ No outcomes found in PDF")
            except Exception as e:
                st.error(f"❌ {e}")

with tab2:
    st.info("💡 One per line: `LO-1: definition`")
    oc_manual = st.text_area(
        "Paste outcomes",
        height=160,
        placeholder="LO-1: Explains machine learning algorithms\nLO-2: Applies supervised learning\nLO-3: Evaluates model performance",
        label_visibility="collapsed"
    )
    if oc_manual.strip():
        for line in oc_manual.strip().split("\n"):
            if ":" in line:
                no, tanim = line.split(":",1)
                no, tanim = no.strip(), tanim.strip()
                if no and tanim:
                    ocler.append({"no": no, "tanim": tanim})
        if ocler:
            st.success(f"✅ {len(ocler)} outcomes defined")

st.divider()

# BÖLÜM 3: Cevap anahtarı
with st.expander("📝 Answer Key (optional)"):
    anahtar = st.text_input("Answer key (ABCDE... format)", placeholder="ABCDEABCDE...").upper()

st.divider()

# BÖLÜM 4: Rapor oluştur
st.markdown("### 📊 Step 3 — Generate Report")

ready = bool(sorular) and bool(ocler)
if not ready:
    if not sorular: st.warning("⬆ Upload an exam file first")
    if not ocler:   st.warning("⬆ Define learning outcomes first")
else:
    st.success(f"Ready: **{len(sorular)} questions** × **{len(ocler)} outcomes**")
    if st.button("🚀 Map with AI & Download Excel", type="primary", use_container_width=True):
        try:
            excel_bytes, mapped = match_and_build_excel(sorular, ocler, anahtar)
            st.success(f"✅ Done! {mapped}/{len(sorular)} questions mapped.")
            st.download_button(
                "⬇ Download Excel Report (.xlsx)",
                excel_bytes,
                f"{exam_filename}-Outcome-Report.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
        except Exception as e:
            st.error(f"❌ Error: {e}")
