import streamlit as st
import json
import re
import io
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Exam Outcome Mapper",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="expanded"
)

# ── SESSION STATE ─────────────────────────────────────────────
for k, v in {
    "logged_in": False,
    "lang": "TR",
    "dark": False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── DİL ──────────────────────────────────────────────────────
TR = {
    "title": "Sınav ÖÇ Eşleştirme",
    "subtitle": "Kapadokya Üniversitesi — Akreditasyon Rapor Sistemi",
    "login": "Giriş Yap",
    "username": "Kullanıcı Adı",
    "password": "Şifre",
    "username_ph": "Kullanıcı adınızı girin",
    "password_ph": "Şifrenizi girin",
    "login_btn": "Giriş Yap",
    "login_err": "❌ Hatalı kullanıcı adı veya şifre",
    "start_over": "🔄 Yeniden Başla",
    "logout": "🚪 Çıkış",
    "step1": "📂 Adım 1 — Sınav Dosyası Yükle",
    "step1_sub": "TXT, DOCX veya PDF",
    "step1_ok": "soru bulundu",
    "step1_err": "❌ Soru bulunamadı. Sorular '?' ile bitmelidir.",
    "preview": "Önizleme",
    "step2": "🎯 Adım 2 — Öğrenim Çıktıları",
    "tab_pdf": "📄 PDF'den Çıkar",
    "tab_manual": "✏️ Elle Gir",
    "pdf_info": "Bilgi Paketi PDF yükleyin — ÖÇ'ler otomatik çıkarılır.",
    "pdf_upload": "Bilgi Paketi PDF",
    "pdf_ok": "ÖÇ çıkarıldı",
    "pdf_err": "❌ ÖÇ bulunamadı",
    "manual_info": "💡 Her satıra: `LO-1: tanım`",
    "manual_ph": "LO-1: Makine öğrenmesi algoritmalarını açıklar\nLO-2: Yöntemleri uygular",
    "outcomes_ok": "öğrenim çıktısı tanımlandı",
    "step3": "💯 Adım 3 — Puanlama",
    "total_pts": "Toplam sınav puanı",
    "scoring_type": "Soru puanlama",
    "equal": "Eşit (otomatik)",
    "custom": "Özel (soru bazlı)",
    "enter_scores": "Her soru için puan girin:",
    "total_warn": "⚠️ Toplam:",
    "total_ok": "✅ Toplam:",
    "answer_key": "📝 Cevap Anahtarı (opsiyonel)",
    "answer_ph": "ABCDEABCDE...",
    "step4": "📊 Adım 4 — Rapor Oluştur",
    "ready_equal": "Hazır:",
    "ready_pts": "puan/soru",
    "ready_custom": "özel puanlama",
    "questions": "soru",
    "outcomes": "çıktı",
    "map_btn": "🚀 AI ile Eşleştir ve Excel Oluştur",
    "mapped_ok": "soru eşleştirildi",
    "multi_match": "birden fazla ÖÇ ile eşleşti",
    "download": "⬇ Excel Raporu İndir (.xlsx)",
    "no_exam": "⬆ Önce sınav dosyası yükleyin",
    "no_oc": "⬆ Önce öğrenim çıktılarını tanımlayın",
    "extracting": "Sorular çıkarılıyor...",
    "oc_extracting": "ÖÇ'ler çıkarılıyor...",
    "mapping": "Eşleştirme yapılıyor...",
    "batch": "Batch",
}

EN = {
    "title": "Exam Outcome Mapper",
    "subtitle": "Kapadokya University — Accreditation Report System",
    "login": "Login",
    "username": "Username",
    "password": "Password",
    "username_ph": "Enter username",
    "password_ph": "Enter password",
    "login_btn": "Login",
    "login_err": "❌ Invalid username or password",
    "start_over": "🔄 Start Over",
    "logout": "🚪 Logout",
    "step1": "📂 Step 1 — Upload Exam File",
    "step1_sub": "TXT, DOCX or PDF",
    "step1_ok": "questions found",
    "step1_err": "❌ No questions found. Questions must end with '?'",
    "preview": "Preview",
    "step2": "🎯 Step 2 — Learning Outcomes",
    "tab_pdf": "📄 Extract from PDF",
    "tab_manual": "✏️ Enter Manually",
    "pdf_info": "Upload Bilgi Paketi PDF — outcomes extracted automatically.",
    "pdf_upload": "Bilgi Paketi PDF",
    "pdf_ok": "outcomes extracted",
    "pdf_err": "❌ No outcomes found",
    "manual_info": "💡 One per line: `LO-1: definition`",
    "manual_ph": "LO-1: Explains ML algorithms\nLO-2: Applies methods",
    "outcomes_ok": "outcomes defined",
    "step3": "💯 Step 3 — Scoring",
    "total_pts": "Total exam points",
    "scoring_type": "Question scoring",
    "equal": "Equal (auto)",
    "custom": "Custom per question",
    "enter_scores": "Enter score for each question:",
    "total_warn": "⚠️ Total:",
    "total_ok": "✅ Total:",
    "answer_key": "📝 Answer Key (optional)",
    "answer_ph": "ABCDEABCDE...",
    "step4": "📊 Step 4 — Generate Report",
    "ready_equal": "Ready:",
    "ready_pts": "pts each",
    "ready_custom": "custom scoring",
    "questions": "questions",
    "outcomes": "outcomes",
    "map_btn": "🚀 Map with AI & Generate Excel",
    "mapped_ok": "questions mapped",
    "multi_match": "matched to multiple outcomes",
    "download": "⬇ Download Excel Report (.xlsx)",
    "no_exam": "⬆ Upload an exam file first",
    "no_oc": "⬆ Define learning outcomes first",
    "extracting": "Extracting questions...",
    "oc_extracting": "Extracting outcomes...",
    "mapping": "Mapping with AI...",
    "batch": "Batch",
}

def t(key):
    d = TR if st.session_state.lang == "TR" else EN
    return d.get(key, key)

# ── TEMA CSS ──────────────────────────────────────────────────
def get_css(dark):
    common = """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif !important; }
    .block-container { padding-top: 0.5rem !important; max-width: 780px !important; }
    /* Streamlit üst header'ı gizle */
    header[data-testid="stHeader"] { display: none !important; }
    #MainMenu { display: none !important; }
    footer { display: none !important; }
    /* Butonlar */
    .stButton > button { border-radius: 8px !important; font-weight: 600 !important; transition: all 0.2s !important; }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #7c6af7, #6d28d9) !important;
        border: none !important; color: white !important;
        box-shadow: 0 4px 15px rgba(124,106,247,0.35) !important;
    }
    .stButton > button[kind="primary"]:hover { transform: translateY(-2px) !important; }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #7c6af7, #6d28d9) !important;
        color: white !important; border: none !important;
        border-radius: 8px !important; font-weight: 600 !important;
        box-shadow: 0 4px 15px rgba(124,106,247,0.35) !important;
        width: 100% !important; padding: 12px !important;
    }
    /* App header */
    .app-header {
        padding: 28px 32px; border-radius: 16px;
        margin-bottom: 28px; text-align: center;
    }
    .app-header h1 { margin:0; font-size:1.6rem; font-weight:800; letter-spacing:-0.5px; }
    .app-header p  { margin:6px 0 0; font-size:0.82rem; }
    /* Section title */
    .section-title {
        font-size:0.7rem; font-weight:700; letter-spacing:2.5px;
        text-transform:uppercase; margin-bottom:14px;
        display:flex; align-items:center; gap:10px;
    }
    .section-title::after { content:''; flex:1; height:1px; }
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { border-radius: 10px !important; padding: 4px !important; }
    .stTabs [data-baseweb="tab"] { border-radius: 8px !important; }
    </style>
    """

    if dark:
        theme = """
    <style>
    .app-header {
        background: linear-gradient(135deg, #1e1b4b 0%, #312e81 60%, #4c1d95 100%);
        box-shadow: 0 8px 32px rgba(124,106,247,0.25);
        border: 1px solid rgba(124,106,247,0.2);
    }
    .app-header h1 { color: #e0e7ff !important; }
    .app-header p  { color: #a5b4fc !important; }
    .section-title { color: #a78bfa; }
    .section-title::after { background: #2d3250; }
    </style>
    """
    else:
        theme = """
    <style>
    .app-header {
        background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 60%, #9333ea 100%);
        box-shadow: 0 8px 32px rgba(109,92,231,0.28);
    }
    .app-header h1 { color: white !important; }
    .app-header p  { color: rgba(255,255,255,0.8) !important; }
    .section-title { color: #6d5ce7; }
    .section-title::after { background: #e2e0f0; }
    .stTabs [data-baseweb="tab-list"] { background: #f1f0ff !important; }
    .stTabs [aria-selected="true"] { background: white !important; box-shadow: 0 2px 8px rgba(0,0,0,0.08) !important; }
    </style>
    """
    return common + theme
st.markdown(get_css(st.session_state.dark), unsafe_allow_html=True)

# ── LOGIN ─────────────────────────────────────────────────────
if not st.session_state.logged_in:
    st.markdown(f"""<div class='app-header'>
        <h1>📋 {t('title')}</h1>
        <p>{t('subtitle')}</p>
    </div>""", unsafe_allow_html=True)

    col_l, col_m, col_r = st.columns([1,2,1])
    with col_m:
        st.markdown(f"### 🔐 {t('login')}")
        username = st.text_input(t("username"), placeholder=t("username_ph"))
        password = st.text_input(t("password"), type="password", placeholder=t("password_ph"))
        if st.button(t("login_btn"), type="primary", use_container_width=True):
            if username == st.secrets.get("APP_USERNAME","admin") and password == st.secrets.get("APP_PASSWORD","admin123"):
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error(t("login_err"))
    st.stop()

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"### 📋 {t('title')}")
    st.markdown("---")

    # Dil & Tema
    c1, c2 = st.columns(2)
    with c1:
        lang = st.radio("🌐", ["TR","EN"], index=0 if st.session_state.lang=="TR" else 1, horizontal=True, label_visibility="collapsed")
        if lang != st.session_state.lang:
            st.session_state.lang = lang
            st.rerun()
    with c2:
        dark = st.toggle("🌙", value=st.session_state.dark)
        if dark != st.session_state.dark:
            st.session_state.dark = dark
            st.rerun()

    st.markdown("---")
    st.markdown(f"👤 **{st.secrets.get('APP_USERNAME','admin')}**")
    if st.button(t("start_over"), use_container_width=True):
        st.rerun()
    if st.button(t("logout"), use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()
    st.markdown("---")
    st.markdown("<small style='color:#888'>Kapadokya University<br>🤖 DeepSeek AI</small>", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────────────────────
st.markdown(f"""<div class='app-header'>
    <h1>📋 {t('title')}</h1>
    <p>{t('subtitle')}</p>
</div>""", unsafe_allow_html=True)

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
    pattern1 = r'(\d{1,3}[\.\)]\s*)([\s\S]*?\?)'
    for idx, (_, soru) in enumerate(re.findall(pattern1, text), 1):
        soru = re.sub(r'\s+', ' ', soru.strip())
        if len(soru) > 8:
            sorular.append({"no": idx, "text": soru})
    if sorular:
        return sorular
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
    for idx, line in enumerate(text.split('\n'), 1):
        line = line.strip()
        if line.endswith('?') and len(line) > 8:
            sorular.append({"no": idx, "text": line})
    return sorular

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

def auto_match(sorular, ocler):
    oc_listesi = "\n".join([f"- {o['no']}: {o['tanim']}" for o in ocler])
    result = {}
    n = len(sorular)
    chunks = (n + 19) // 20
    progress = st.progress(0, text=f"{t('mapping')} (1/{chunks})")

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
            st.warning(f"{t('batch')} {ci+1} error: {e}")
        progress.progress((ci+1)/chunks, text=f"{t('mapping')} ({ci+1}/{chunks})")

    progress.empty()
    return result

def build_excel(sorular, ocler, eslestirmeler, anahtar, puan_esit, toplam_puan, ozel_puanlar):
    wb = Workbook()
    PURPLE="5B4FCF"; TEAL="0F766E"; GREEN_L="F0FDF4"
    GRAY="F8F7FF"; WHITE="FFFFFF"; DARK="1A1640"

    def hstyle(cell, bg=PURPLE):
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def add_borders(ws, r1, r2, c1, c2):
        t = Side(style="thin", color="E2E0F0")
        for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
            for cell in row:
                cell.border = Border(left=t, right=t, top=t, bottom=t)

    oc_no_to_idx = {o["no"]: i+1 for i, o in enumerate(ocler)}

    def get_puan(soru_no):
        if puan_esit:
            return round(toplam_puan / len(sorular), 2)
        return ozel_puanlar.get(soru_no, round(toplam_puan / len(sorular), 2))

    max_oc = max((len(eslestirmeler.get(s["no"],{}).get("outcomes",[])) for s in sorular), default=1)
    max_oc = max(max_oc, 1)

    # ─── SAYFA 1: Question-LO Mapping ───────────────────────
    ws1 = wb.active
    ws1.title = "Question-LO Mapping"
    ws1.row_dimensions[1].height = 36

    headers1 = ["Q NO", "SCORE", "QUESTION"]
    for i in range(max_oc):
        headers1.append(f"DÇ Sıra {i+1}")
        headers1.append(f"Etki Oran {i+1}")
    headers1 += ["DIFFICULTY", "ANSWER KEY"]

    for col, h in enumerate(headers1, 1):
        hstyle(ws1.cell(row=1, column=col, value=h))

    for ri, s in enumerate(sorular, 2):
        esl = eslestirmeler.get(s["no"], {})
        outcomes = esl.get("outcomes", [])
        diff = esl.get("zorluk","")
        key = anahtar[s["no"]-1] if anahtar and s["no"]-1 < len(anahtar) else "-"
        puan = get_puan(s["no"])

        ws1.cell(ri,1,s["no"]).alignment = Alignment(horizontal="center")
        ws1.cell(ri,2,puan).alignment    = Alignment(horizontal="center")
        ws1.cell(ri,3,s["text"]).alignment = Alignment(wrap_text=True, vertical="top")

        for i, outcome in enumerate(outcomes):
            lo_idx = oc_no_to_idx.get(outcome.get("lo",""), "")
            pct    = outcome.get("pct", 100)
            cb = 4 + i*2
            ws1.cell(ri, cb,   lo_idx).alignment = Alignment(horizontal="center")
            ws1.cell(ri, cb+1, pct).alignment    = Alignment(horizontal="center")

        diff_col = 4 + max_oc*2
        ws1.cell(ri, diff_col,   diff).alignment = Alignment(horizontal="center")
        ws1.cell(ri, diff_col+1, key).alignment  = Alignment(horizontal="center")

        if outcomes:
            for c in range(1, len(headers1)+1):
                ws1.cell(ri,c).fill = PatternFill("solid", fgColor="EDE9FE")
        elif ri%2==0:
            for c in range(1, len(headers1)+1):
                ws1.cell(ri,c).fill = PatternFill("solid", fgColor=GRAY)

    ws1.column_dimensions["A"].width = 8
    ws1.column_dimensions["B"].width = 8
    ws1.column_dimensions["C"].width = 58
    for i in range(max_oc):
        ws1.column_dimensions[get_column_letter(4+i*2)].width   = 10
        ws1.column_dimensions[get_column_letter(4+i*2+1)].width = 12
    ws1.column_dimensions[get_column_letter(4+max_oc*2)].width   = 12
    ws1.column_dimensions[get_column_letter(4+max_oc*2+1)].width = 12
    ws1.freeze_panes = "A2"
    add_borders(ws1, 1, len(sorular)+1, 1, len(headers1))

    # ─── SAYFA 2: LO Summary ────────────────────────────────
    ws2 = wb.create_sheet("LO Summary")
    ws2.row_dimensions[1].height = 36
    for col, h in enumerate(["LO NO","LO DEFINITION","# QUESTIONS","QUESTION NUMBERS"],1):
        hstyle(ws2.cell(row=1, column=col, value=h), bg=TEAL)

    for ri, oc in enumerate(ocler, 2):
        matched = [s for s in sorular if any(o["lo"]==oc["no"] for o in eslestirmeler.get(s["no"],{}).get("outcomes",[]))]
        q_nums = ", ".join(str(s["no"]) for s in matched)
        ws2.cell(ri,1,oc["no"]).alignment = Alignment(horizontal="center")
        ws2.cell(ri,2,oc["tanim"]).alignment = Alignment(wrap_text=True)
        ws2.cell(ri,3,len(matched)).alignment = Alignment(horizontal="center")
        ws2.cell(ri,4,q_nums).alignment = Alignment(wrap_text=True)
        if ri%2==0:
            for c in [1,2,3,4]: ws2.cell(ri,c).fill = PatternFill("solid",fgColor="F0FDFA")

    for col,w in zip("ABCD",[10,52,14,40]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"
    add_borders(ws2, 1, len(ocler)+1, 1, 4)

    # ─── SAYFA 3: Proliz Format ─────────────────────────────
    ws3 = wb.create_sheet("Proliz Format")
    ws3.row_dimensions[1].height = 36

    headers3 = ["Soru No", "Soru Puan", "Soru Metni"]
    for i in range(max_oc):
        headers3.append(f"DÇ Sıra {i+1}")
        headers3.append(f"Etki Oran {i+1}")

    for col, h in enumerate(headers3, 1):
        hstyle(ws3.cell(row=1, column=col, value=h))

    for ri, s in enumerate(sorular, 2):
        esl = eslestirmeler.get(s["no"], {})
        outcomes = esl.get("outcomes", [])
        puan = get_puan(s["no"])

        ws3.cell(ri,1,s["no"]).alignment = Alignment(horizontal="center")
        ws3.cell(ri,2,puan).alignment    = Alignment(horizontal="center")
        ws3.cell(ri,3,s["text"]).alignment = Alignment(wrap_text=True, vertical="top")

        for i, outcome in enumerate(outcomes):
            lo_idx = oc_no_to_idx.get(outcome.get("lo",""), "")
            pct    = outcome.get("pct", 100)
            cb = 4 + i*2
            ws3.cell(ri, cb,   lo_idx).alignment = Alignment(horizontal="center")
            ws3.cell(ri, cb+1, pct).alignment    = Alignment(horizontal="center")

        if outcomes:
            for c in range(1, len(headers3)+1):
                ws3.cell(ri,c).fill = PatternFill("solid", fgColor="EDE9FE")
        elif ri%2==0:
            for c in range(1, len(headers3)+1):
                ws3.cell(ri,c).fill = PatternFill("solid", fgColor=GRAY)

    ws3.column_dimensions["A"].width = 9
    ws3.column_dimensions["B"].width = 10
    ws3.column_dimensions["C"].width = 60
    for i in range(max_oc):
        ws3.column_dimensions[get_column_letter(4+i*2)].width   = 10
        ws3.column_dimensions[get_column_letter(4+i*2+1)].width = 12
    ws3.freeze_panes = "A2"
    add_borders(ws3, 1, len(sorular)+1, 1, len(headers3))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════

# ADIM 1
st.markdown(f"<div class='section-title'>{t('step1')}</div>", unsafe_allow_html=True)
uploaded_exam = st.file_uploader(t("step1_sub"), type=["txt","pdf","docx","doc"], key="exam", label_visibility="collapsed")

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
            st.success(f"✅ {len(sorular)} {t('step1_ok')}")
            with st.expander(f"👁 {t('preview')}", expanded=False):
                for s in sorular:
                    st.markdown(f"**{s['no']}.** {s['text'][:120]}...")
        else:
            st.error(t("step1_err"))
    except Exception as e:
        st.error(f"❌ {e}")

st.divider()

# ADIM 2
st.markdown(f"<div class='section-title'>{t('step2')}</div>", unsafe_allow_html=True)
tab1, tab2 = st.tabs([t("tab_pdf"), t("tab_manual")])

ocler = []
with tab1:
    st.info(t("pdf_info"))
    uploaded_bp = st.file_uploader(t("pdf_upload"), type=["pdf"], key="bilgi")
    if uploaded_bp:
        with st.spinner(t("oc_extracting")):
            try:
                extracted = extract_oc_from_pdf(uploaded_bp.read())
                if extracted:
                    st.success(f"✅ {len(extracted)} {t('pdf_ok')}")
                    for oc in extracted:
                        st.markdown(f"**{oc['no']}** — {oc['tanim'][:80]}...")
                    ocler = extracted
                else:
                    st.error(t("pdf_err"))
            except Exception as e:
                st.error(f"❌ {e}")

with tab2:
    st.info(t("manual_info"))
    oc_manual = st.text_area("", height=160, placeholder=t("manual_ph"))
    if oc_manual.strip():
        for line in oc_manual.strip().split("\n"):
            if ":" in line:
                no, tanim = line.split(":",1)
                no, tanim = no.strip(), tanim.strip()
                if no and tanim:
                    ocler.append({"no":no,"tanim":tanim})
        if ocler:
            st.success(f"✅ {len(ocler)} {t('outcomes_ok')}")

st.divider()

# ADIM 3
st.markdown(f"<div class='section-title'>{t('step3')}</div>", unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    toplam_puan = st.number_input(t("total_pts"), min_value=1, max_value=1000, value=100)
with col2:
    puan_tipi = st.radio(t("scoring_type"), [t("equal"), t("custom")], horizontal=True)

puan_esit = puan_tipi == t("equal")
ozel_puanlar = {}

if not puan_esit and sorular:
    st.markdown(f"**{t('enter_scores')}**")
    cols = st.columns(5)
    for i, s in enumerate(sorular):
        with cols[i % 5]:
            ozel_puanlar[s["no"]] = st.number_input(
                f"Q{s['no']}", min_value=0.0, max_value=float(toplam_puan),
                value=round(toplam_puan/len(sorular),1),
                key=f"puan_{s['no']}"
            )
    total_check = sum(ozel_puanlar.values())
    if abs(total_check - toplam_puan) > 0.5:
        st.warning(f"{t('total_warn')} {total_check} (= {toplam_puan})")
    else:
        st.success(f"{t('total_ok')} {total_check}")

st.divider()

# ADIM 4: Cevap anahtarı
with st.expander(t("answer_key")):
    anahtar = st.text_input("", placeholder=t("answer_ph")).upper()

st.divider()

# ADIM 5: Rapor
st.markdown(f"<div class='section-title'>{t('step4')}</div>", unsafe_allow_html=True)

ready = bool(sorular) and bool(ocler)
if not ready:
    if not sorular: st.warning(t("no_exam"))
    if not ocler:   st.warning(t("no_oc"))
else:
    if puan_esit:
        puan_per_q = round(toplam_puan / len(sorular), 2)
        st.success(f"{t('ready_equal')} **{len(sorular)} {t('questions')}** × **{len(ocler)} {t('outcomes')}** — {puan_per_q} {t('ready_pts')}")
    else:
        st.success(f"{t('ready_equal')} **{len(sorular)} {t('questions')}** × **{len(ocler)} {t('outcomes')}** — {t('ready_custom')}")

    if st.button(t("map_btn"), type="primary", use_container_width=True):
        try:
            eslestirmeler = auto_match(sorular, ocler)
            mapped = sum(1 for v in eslestirmeler.values() if v.get("outcomes"))
            multi  = sum(1 for v in eslestirmeler.values() if len(v.get("outcomes",[])) > 1)
            st.success(f"✅ {mapped}/{len(sorular)} {t('mapped_ok')} — {multi} {t('multi_match')}")

            excel_bytes = build_excel(sorular, ocler, eslestirmeler, anahtar, puan_esit, toplam_puan, ozel_puanlar)
            st.download_button(
                t("download"), excel_bytes,
                f"{exam_filename}-Outcome-Report.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary"
            )
        except Exception as e:
            st.error(f"❌ {e}")
