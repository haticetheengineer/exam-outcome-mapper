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
    background: white; border: 1px solid #e0e0e0; border-radius: 12px;
    padding: 32px; max-width: 400px; margin: 60px auto;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}
</style>
""", unsafe_allow_html=True)

# ── LOGIN ─────────────────────────────────────────────────────
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown("""<div class='header'>
        <h1>📋 Exam Outcome Mapper</h1>
        <p>Kapadokya University — Accreditation Report System</p>
    </div>""", unsafe_allow_html=True)
    with st.container():
        st.markdown("### 🔐 Login")
        username = st.text_input("Username", placeholder="Enter username")
        password = st.text_input("Password", type="password", placeholder="Enter password")
        if st.button("Login", type="primary", use_container_width=True):
            if username == st.secrets.get("APP_USERNAME","admin") and password == st.secrets.get("APP_PASSWORD","admin123"):
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("❌ Invalid username or password")
    st.stop()

# ── HEADER ────────────────────────────────────────────────────
st.markdown("""<div class='header'>
    <h1>📋 Exam Outcome Mapper</h1>
    <p>Kapadokya University — Accreditation Report System</p>
</div>""", unsafe_allow_html=True)

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Exam Outcome Mapper")
    st.markdown("---")
    st.markdown(f"👤 **{st.secrets.get('APP_USERNAME','admin')}**")
    if st.button("🔄 Start Over", use_container_width=True):
        st.rerun()
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()
    st.markdown("---")
    st.markdown("<small style='color:#aaa'>Kapadokya University<br>🤖 DeepSeek AI</small>", unsafe_allow_html=True)

# ── DEEPSEEK ──────────────────────────────────────────────────
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"

def deepseek_chat(messages, max_tokens=4096):
    resp = requests.post(
        DEEPSEEK_URL,
        headers={"Authorization": f"Bearer {st.secrets['DEEPSEEK_API_KEY']}", "Content-Type": "application/json"},
        json={"model": "deepseek-chat", "messages": messages, "max_tokens": max_tokens},
        timeout=60
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"].strip()

# ── SORU ÇIKARMA ──────────────────────────────────────────────
def extract_questions(text):
    sorular = []

    # Format 1: Numaralı + "?" (normal döküman)
    pattern1 = r'(\d{1,3}[\.\)]\s*)([\s\S]*?\?)'
    for idx, (_, soru) in enumerate(re.findall(pattern1, text), 1):
        soru = re.sub(r'\s+', ' ', soru.strip())
        if len(soru) > 8:
            sorular.append({"no": idx, "text": soru})
    if sorular:
        return sorular

    # Format 2: "Soru X." tablo formatı
    pattern2 = r'Soru\s+(\d+)\.\s+(.*?)(?=Soru\s+\d+\.|$)'
    for no_str, metin in re.findall(pattern2, text, re.DOTALL):
        metin = re.sub(r'\n\s*[A-E]\)\s.*', '', metin)
        metin = re.sub(r'Cevap:.*', '', metin, flags=re.DOTALL)
        metin = re.sub(r'\|', '', metin)
        metin = re.sub(r'\s+', ' ', metin.strip())
        if len(metin) > 8:
            sorular.append({"no": int(no_str), "text": metin})
    if sorular:
        return sorular

    # Format 3: Satır sonu "?"
    for idx, line in enumerate(text.split('\n'), 1):
        line = line.strip()
        if line.endswith('?') and len(line) > 8:
            sorular.append({"no": idx, "text": line})
    return sorular

# ── ÖÇ PDF ÇIKARMA ───────────────────────────────────────────
def extract_oc_from_pdf(pdf_bytes):
    import pypdf
    text = "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(io.BytesIO(pdf_bytes)).pages)
    prompt = f"""Extract all learning outcomes (Dersin Öğrenme Çıktıları) from this Turkish university course info PDF.
Return ONLY JSON:
{{"outcomes":[{{"no":"LO-1","definition":"..."}},{{"no":"LO-2","definition":"..."}}]}}
TEXT: {text[:8000]}"""
    raw = deepseek_chat([{"role":"user","content":prompt}], max_tokens=2048)
    raw = re.sub(r'```json|```','',raw).strip()
    m = re.search(r'\{[\s\S]*\}', raw)
    if m:
        data = json.loads(m.group())
        return [{"no": o["no"], "tanim": o["definition"]} for o in data.get("outcomes", [])]
    return []

# ── OTOMATİK EŞLEŞTİRME (çoklu ÖÇ destekli) ─────────────────
def auto_match(sorular, ocler):
    """Her soruyu 1 veya birden fazla ÖÇ ile eşleştir, yüzdelerini belirle"""
    oc_listesi = "\n".join([f"- {o['no']}: {o['tanim']}" for o in ocler])
    result = {}
    n = len(sorular)
    chunks = (n + 19) // 20
    progress = st.progress(0, text="AI mapping...")

    for ci, i in enumerate(range(0, n, 20)):
        chunk = sorular[i:i+20]
        soru_listesi = "\n".join([f"{s['no']}. {s['text']}" for s in chunk])

        prompt = f"""You are an academic assessment expert. Match each exam question to learning outcomes.
A question CAN match multiple outcomes if it covers multiple topics.

LEARNING OUTCOMES:
{oc_listesi}

EXAM QUESTIONS:
{soru_listesi}

Rules:
- Match to 1 outcome if clearly focused on one topic
- Match to 2-3 outcomes if the question spans multiple topics
- Percentages must sum to 100 for each question
- Assign difficulty: Easy, Medium, Hard

Return ONLY valid JSON:
{{"matches":[
  {{"q":1,"outcomes":[{{"lo":"LO-3","pct":100}}],"difficulty":"Medium"}},
  {{"q":2,"outcomes":[{{"lo":"LO-1","pct":60}},{{"lo":"LO-2","pct":40}}],"difficulty":"Easy"}}
]}}"""

        try:
            raw = deepseek_chat([{"role":"user","content":prompt}], max_tokens=2048)
            raw = re.sub(r'```json|```','',raw).strip()
            m = re.search(r'\{[\s\S]*\}', raw)
            if m:
                data = json.loads(m.group())
                for item in data.get("matches", []):
                    q_no = int(item.get("q", 0))
                    outcomes = item.get("outcomes", [])
                    diff = str(item.get("difficulty", "Medium"))
                    if q_no > 0 and outcomes:
                        result[q_no] = {"outcomes": outcomes, "zorluk": diff}
        except Exception as e:
            st.warning(f"Batch {ci+1} error: {e}")

        progress.progress((ci+1)/chunks, text=f"Batch {ci+1}/{chunks}...")

    progress.empty()
    return result

# ── EXCEL OLUŞTURMA ───────────────────────────────────────────
def build_excel(sorular, ocler, eslestirmeler, anahtar, puan_esit, toplam_puan, ozel_puanlar):
    wb = Workbook()
    BLUE="1A3A6B"; GREEN="1E6B3A"; GREEN_L="E8F5ED"
    RED="C0392B"; GRAY="F5F5F5"; WHITE="FFFFFF"; YELLOW="FFF9C4"

    def hstyle(cell, bg=BLUE):
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def borders(ws, r1, r2, c1, c2):
        t = Side(style="thin", color="CCCCCC")
        for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
            for cell in row:
                cell.border = Border(left=t, right=t, top=t, bottom=t)

    oc_map = {o["no"]: o["tanim"] for o in ocler}
    oc_no_to_idx = {o["no"]: i+1 for i, o in enumerate(ocler)}  # LO-1 → 1, LO-2 → 2

    # Soru puanı hesapla
    n = len(sorular)
    def get_puan(soru_no):
        if puan_esit:
            return round(toplam_puan / n, 2)
        return ozel_puanlar.get(soru_no, round(toplam_puan / n, 2))

    # ─── SAYFA 1: Question-LO Mapping ───────────────────────
    ws1 = wb.active
    ws1.title = "Question-LO Mapping"
    ws1.row_dimensions[1].height = 32
    for col, h in enumerate(["Q NO","QUESTION","ANSWER KEY","SCORE","LO NO","LO WEIGHT%","DIFFICULTY"], 1):
        hstyle(ws1.cell(row=1, column=col, value=h))

    for ri, s in enumerate(sorular, 2):
        esl = eslestirmeler.get(s["no"], {})
        outcomes = esl.get("outcomes", [])
        lo_no   = ", ".join(str(oc_no_to_idx.get(o["lo"], o["lo"])) for o in outcomes) if outcomes else ""
        lo_pct  = ", ".join(str(o["pct"])+"%" for o in outcomes) if outcomes else ""
        diff = esl.get("zorluk","")
        key = anahtar[s["no"]-1] if anahtar and s["no"]-1 < len(anahtar) else "-"
        puan = get_puan(s["no"])

        ws1.cell(ri,1,s["no"]).alignment = Alignment(horizontal="center")
        ws1.cell(ri,2,s["text"]).alignment = Alignment(wrap_text=True, vertical="top")
        ws1.cell(ri,3,key).alignment = Alignment(horizontal="center")
        ws1.cell(ri,4,puan).alignment = Alignment(horizontal="center")
        ws1.cell(ri,5,lo_no).alignment  = Alignment(horizontal="center", wrap_text=True)
        ws1.cell(ri,6,lo_pct).alignment = Alignment(horizontal="center")
        ws1.cell(ri,7,diff).alignment = Alignment(horizontal="center")

        if outcomes:
            for c in range(1,8):
                ws1.cell(ri,c).fill = PatternFill("solid", fgColor=GREEN_L)
        elif ri%2==0:
            for c in range(1,8):
                ws1.cell(ri,c).fill = PatternFill("solid", fgColor=GRAY)

    for col, w in zip("ABCDEFG",[8,58,12,8,12,40,12]):
        ws1.column_dimensions[col].width = w
    ws1.freeze_panes = "A2"
    borders(ws1,1,len(sorular)+1,1,7)

    # ─── SAYFA 2: LO Summary ────────────────────────────────
    ws2 = wb.create_sheet("LO Summary")
    ws2.row_dimensions[1].height = 32
    for col, h in enumerate(["LO NO","LO DEFINITION","# QUESTIONS","QUESTION NUMBERS"],1):
        hstyle(ws2.cell(row=1, column=col, value=h))

    for ri, oc in enumerate(ocler, 2):
        matched = [s for s in sorular if any(o["lo"]==oc["no"] for o in eslestirmeler.get(s["no"],{}).get("outcomes",[]))]
        q_nums = ", ".join(str(s["no"]) for s in matched)
        ws2.cell(ri,1,oc["no"]).alignment = Alignment(horizontal="center")
        ws2.cell(ri,2,oc["tanim"]).alignment = Alignment(wrap_text=True)
        ws2.cell(ri,3,len(matched)).alignment = Alignment(horizontal="center")
        ws2.cell(ri,4,q_nums).alignment = Alignment(wrap_text=True)
        if ri%2==0:
            for c in [1,2,3,4]: ws2.cell(ri,c).fill = PatternFill("solid",fgColor=GRAY)

    for col,w in zip("ABCD",[10,52,14,40]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"
    borders(ws2,1,len(ocler)+1,1,4)

    # ─── SAYFA 3: Proliz Formatı ────────────────────────────
    ws3 = wb.create_sheet("Proliz Format")
    ws3.row_dimensions[1].height = 40

    # Maksimum kaç ÖÇ eşleşmesi var?
    max_oc = max((len(eslestirmeler.get(s["no"],{}).get("outcomes",[])) for s in sorular), default=1)
    max_oc = max(max_oc, 1)

    # Header
    headers = ["Soru No", "Soru Puan", "Soru Metni"]
    for i in range(max_oc):
        headers.append(f"DÇ Sıra {i+1}")
        headers.append(f"Etki Oran {i+1}")

    for col, h in enumerate(headers, 1):
        c = ws3.cell(row=1, column=col, value=h)
        hstyle(c, bg="1A3A6B")

    for ri, s in enumerate(sorular, 2):
        esl = eslestirmeler.get(s["no"], {})
        outcomes = esl.get("outcomes", [])
        puan = get_puan(s["no"])

        ws3.cell(ri,1,s["no"]).alignment = Alignment(horizontal="center")
        ws3.cell(ri,2,puan).alignment = Alignment(horizontal="center")
        ws3.cell(ri,3,s["text"]).alignment = Alignment(wrap_text=True, vertical="top")

        for i, outcome in enumerate(outcomes):
            lo_no = outcome.get("lo","")
            pct   = outcome.get("pct", 100)
            lo_idx = oc_no_to_idx.get(lo_no, "")  # LO-4 → 4
            col_base = 4 + i*2
            ws3.cell(ri, col_base,   lo_idx).alignment = Alignment(horizontal="center")
            ws3.cell(ri, col_base+1, pct).alignment    = Alignment(horizontal="center")

        # Renk
        if outcomes:
            for c in range(1, len(headers)+1):
                ws3.cell(ri,c).fill = PatternFill("solid", fgColor=GREEN_L)
        elif ri%2==0:
            for c in range(1, len(headers)+1):
                ws3.cell(ri,c).fill = PatternFill("solid", fgColor=GRAY)

    ws3.column_dimensions["A"].width = 9
    ws3.column_dimensions["B"].width = 10
    ws3.column_dimensions["C"].width = 60
    for i in range(max_oc):
        col_letter_dc  = chr(ord("D") + i*2)
        col_letter_eti = chr(ord("E") + i*2)
        ws3.column_dimensions[col_letter_dc].width  = 10
        ws3.column_dimensions[col_letter_eti].width = 12

    ws3.freeze_panes = "A2"
    borders(ws3, 1, len(sorular)+1, 1, len(headers))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════

# ADIM 1: Sınav Dosyası
st.markdown("### 📂 Step 1 — Upload Exam File")
uploaded_exam = st.file_uploader("TXT, DOCX or PDF", type=["txt","pdf","docx","doc"], key="exam")

sorular = []
exam_filename = "Exam-Outcome-Report"
if uploaded_exam:
    exam_filename = uploaded_exam.name.rsplit(".",1)[0]
    ext = uploaded_exam.name.split(".")[-1].lower()
    raw = uploaded_exam.read()
    try:
        if ext == "txt":
            text = raw.decode("utf-8", errors="ignore")
            sorular = extract_questions(text)
        elif ext in ["docx","doc"]:
            from docx import Document
            doc = Document(io.BytesIO(raw))
            para_text = "\n".join(p.text for p in doc.paragraphs)
            sorular = extract_questions(para_text)
            if not sorular:
                seen = set()
                table_text = ""
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            ct = cell.text.strip()
                            if ct and ct not in seen:
                                seen.add(ct)
                                table_text += ct + "\n"
                sorular = extract_questions(table_text)
        elif ext == "pdf":
            import pypdf
            text = "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(io.BytesIO(raw)).pages)
            sorular = extract_questions(text)

        if sorular:
            st.success(f"✅ {len(sorular)} questions found")
            with st.expander("Preview", expanded=False):
                for s in sorular:
                    st.markdown(f"**{s['no']}.** {s['text'][:120]}...")
        else:
            st.error("❌ No questions found")
    except Exception as e:
        st.error(f"❌ {e}")

st.divider()

# ADIM 2: ÖÇ
st.markdown("### 🎯 Step 2 — Learning Outcomes")
tab1, tab2 = st.tabs(["📄 Extract from PDF", "✏️ Enter Manually"])

ocler = []
with tab1:
    st.info("Upload Bilgi Paketi PDF — outcomes extracted automatically.")
    uploaded_bp = st.file_uploader("Bilgi Paketi PDF", type=["pdf"], key="bilgi")
    if uploaded_bp:
        with st.spinner("Extracting outcomes..."):
            try:
                extracted = extract_oc_from_pdf(uploaded_bp.read())
                if extracted:
                    st.success(f"✅ {len(extracted)} outcomes extracted!")
                    for oc in extracted:
                        st.markdown(f"**{oc['no']}** — {oc['tanim'][:80]}...")
                    ocler = extracted
                else:
                    st.error("❌ No outcomes found")
            except Exception as e:
                st.error(f"❌ {e}")

with tab2:
    st.info("💡 One per line: `LO-1: definition`")
    oc_manual = st.text_area("Paste outcomes", height=160, label_visibility="collapsed",
        placeholder="LO-1: Explains ML algorithms\nLO-2: Applies methods")
    if oc_manual.strip():
        for line in oc_manual.strip().split("\n"):
            if ":" in line:
                no, tanim = line.split(":",1)
                no, tanim = no.strip(), tanim.strip()
                if no and tanim:
                    ocler.append({"no":no,"tanim":tanim})
        if ocler:
            st.success(f"✅ {len(ocler)} outcomes defined")

st.divider()

# ADIM 3: Puan Ayarları
st.markdown("### 💯 Step 3 — Scoring")

col1, col2 = st.columns(2)
with col1:
    toplam_puan = st.number_input("Total exam points", min_value=1, max_value=1000, value=100)
with col2:
    puan_tipi = st.radio("Question scoring", ["Equal (auto)", "Custom per question"], horizontal=True)

puan_esit = puan_tipi == "Equal (auto)"
ozel_puanlar = {}

if not puan_esit and sorular:
    st.markdown("**Enter score for each question:**")
    cols = st.columns(5)
    for i, s in enumerate(sorular):
        with cols[i % 5]:
            ozel_puanlar[s["no"]] = st.number_input(
                f"Q{s['no']}", min_value=0.0, max_value=float(toplam_puan),
                value=round(toplam_puan/len(sorular),1),
                key=f"puan_{s['no']}", label_visibility="visible"
            )
    total_check = sum(ozel_puanlar.values())
    if abs(total_check - toplam_puan) > 0.5:
        st.warning(f"⚠️ Total: {total_check} (should be {toplam_puan})")
    else:
        st.success(f"✅ Total: {total_check}")

st.divider()

# ADIM 4: Cevap anahtarı
with st.expander("📝 Answer Key (optional)"):
    anahtar = st.text_input("Answer key (ABCDE...)", placeholder="ABCDEABCDE...").upper()

st.divider()

# ADIM 5: Rapor
st.markdown("### 📊 Step 4 — Generate Report")

ready = bool(sorular) and bool(ocler)
if not ready:
    if not sorular: st.warning("⬆ Upload an exam file first")
    if not ocler:   st.warning("⬆ Define learning outcomes first")
else:
    if puan_esit:
        puan_per_q = round(toplam_puan / len(sorular), 2)
        st.success(f"Ready: **{len(sorular)} questions** × **{len(ocler)} outcomes** — {puan_per_q} pts each")
    else:
        st.success(f"Ready: **{len(sorular)} questions** × **{len(ocler)} outcomes** — custom scoring")

    if st.button("🚀 Map with AI & Generate Excel", type="primary", use_container_width=True):
        try:
            eslestirmeler = auto_match(sorular, ocler)
            mapped = sum(1 for v in eslestirmeler.values() if v.get("outcomes"))
            multi = sum(1 for v in eslestirmeler.values() if len(v.get("outcomes",[])) > 1)

            st.success(f"✅ {mapped}/{len(sorular)} questions mapped — {multi} matched to multiple outcomes")

            excel_bytes = build_excel(sorular, ocler, eslestirmeler, anahtar, puan_esit, toplam_puan, ozel_puanlar)

            st.download_button(
                "⬇ Download Excel Report (.xlsx)",
                excel_bytes,
                f"{exam_filename}-Outcome-Report.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary"
            )
        except Exception as e:
            st.error(f"❌ {e}")
