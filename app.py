import streamlit as st
import json
import re
import io
import base64
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Exam Outcome Mapper", page_icon="📋", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.step-header {
    background: linear-gradient(135deg, #1a3a6b, #2c5aa0);
    color: white; padding: 16px 20px; border-radius: 10px; margin-bottom: 20px;
}
.step-header h2 { margin: 0; font-size: 1.1rem; }
.step-header p  { margin: 4px 0 0; font-size: 0.8rem; opacity: 0.8; }
.soru-card {
    background: #f8f9ff; border: 1px solid #d0d8ee;
    border-left: 4px solid #1a3a6b; border-radius: 8px;
    padding: 12px 14px; margin-bottom: 10px;
}
.soru-card.matched { border-left-color: #1e6b3a; background: #f0fdf4; }
.soru-no { color: #c0392b; font-weight: 700; font-size: 0.85rem; }
.soru-text { font-size: 0.88rem; line-height: 1.5; margin-top: 4px; }
.oc-badge {
    background: #1a3a6b; color: white; padding: 2px 10px;
    border-radius: 12px; font-size: 0.72rem; font-weight: 600;
}
.auto-banner {
    background: #fffbeb; border: 1px solid #fcd34d;
    border-radius: 8px; padding: 12px 16px;
    font-size: 0.82rem; color: #92400e; margin-bottom: 14px;
}
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ─────────────────────────────────────────────
def init_state():
    defaults = {
        "step": 1, "sorular": [], "ocler": [],
        "eslestirmeler": {}, "anahtar": "", "ogrenciler": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ── DEEPSEEK API ──────────────────────────────────────────────
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"

def deepseek_chat(messages: list, max_tokens=4096) -> str:
    api_key = st.secrets["DEEPSEEK_API_KEY"]
    resp = requests.post(
        DEEPSEEK_URL,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={"model": "deepseek-chat", "messages": messages, "max_tokens": max_tokens},
        timeout=60
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"].strip()

def deepseek_vision(img_bytes: bytes, mime: str, prompt: str) -> str:
    """DeepSeek gorsel desteklemez — pytesseract ile metni cikar, sonra DeepSeek'e gonder"""
    try:
        from PIL import Image
        import pytesseract
        img = Image.open(io.BytesIO(img_bytes))
        text = pytesseract.image_to_string(img, lang="tur+eng")
        if not text.strip():
            raise ValueError("Resimden metin okunamadi")
        full_prompt = f"{prompt}\n\nResimden okunan metin:\n{text[:3000]}"
        return deepseek_chat([{"role": "user", "content": full_prompt}], max_tokens=2000)
    except ImportError:
        raise Exception("Resimden ÖÇ okuma için pytesseract gerekli. Lütfen ÖÇ'leri elle girin.")

# ── OTOMATİK EŞLEŞTİRME ──────────────────────────────────────
def auto_match(sorular: list, ocler: list) -> dict:
    """DeepSeek ile her soruyu en uygun ÖÇ ile eşleştir — 20'şer parça halinde"""
    oc_listesi = "\n".join([f"- {o['no']}: {o['tanim']}" for o in ocler])
    result = {}
    chunk_size = 20

    for i in range(0, len(sorular), chunk_size):
        chunk = sorular[i:i+chunk_size]
        soru_listesi = "\n".join([f"{s['no']}. {s['text']}" for s in chunk])

        prompt = f"""Asagida ogrenim ciktilari (OC) ve sinav sorulari var.
Her soruyu en uygun OC ile eslestir.

OGRENIM CIKTILARI:
{oc_listesi}

SINAV SORULARI:
{soru_listesi}

SADECE JSON dondur, baska hicbir sey yazma:
{{
  "eslestirmeler": [
    {{"soru_no": 1, "oc_no": "LO-1", "zorluk": "Medium"}},
    {{"soru_no": 2, "oc_no": "LO-2", "zorluk": "Easy"}}
  ]
}}"""

        try:
            raw = deepseek_chat([{"role": "user", "content": prompt}], max_tokens=4096)
            raw = re.sub(r'```json|```', '', raw).strip()
            data = json.loads(raw)
            for item in data.get("eslestirmeler", []):
                result[item["soru_no"]] = {
                    "oc_no": item.get("oc_no", ""),
                    "zorluk": item.get("zorluk", "Medium")
                }
        except Exception:
            # Bu chunk basarisiz olursa devam et
            for s in chunk:
                result[s["no"]] = {"oc_no": ocler[0]["no"] if ocler else "", "zorluk": "Medium"}

    return result

# ── YARDIMCI ─────────────────────────────────────────────────
def parse_text(text: str) -> list:
    sorular = []
    pattern = r'(\d{1,3}[\.\)]\s*)([\s\S]*?\?)'
    matches = re.findall(pattern, text)
    for idx, (_, soru) in enumerate(matches, 1):
        soru = re.sub(r'\s+', ' ', soru.strip().replace('\n', ' '))
        if len(soru) > 8:
            sorular.append({"no": idx, "text": soru})
    if not sorular:
        for idx, line in enumerate(text.split('\n'), 1):
            line = line.strip()
            if line.endswith('?') and len(line) > 8:
                sorular.append({"no": idx, "text": line})
    return sorular

def parse_via_deepseek(text: str) -> list:
    prompt = f"""Asagidaki sinav metninden "?" ile biten TUM sorulari cikar.
SADECE JSON dondur:
{{"sorular":[{{"no":1,"text":"Soru metni?"}},{{"no":2,"text":"..."}}]}}

METIN:
{text[:10000]}"""
    raw = deepseek_chat([{"role": "user", "content": prompt}])
    raw = re.sub(r'```json|```', '', raw).strip()
    return json.loads(raw).get("sorular", [])

def read_oc_from_image(img_bytes: bytes, mime: str) -> list:
    prompt = """Bu resimde ogrenim ciktilari listesi var.
Tum OC'leri cikar. SADECE JSON dondur:
{"ocler":[{"no":"LO-1","tanim":"..."},{"no":"LO-2","tanim":"..."}]}"""
    raw = deepseek_vision(img_bytes, mime, prompt)
    raw = re.sub(r'```json|```', '', raw).strip()
    return json.loads(raw).get("ocler", [])

def build_excel(sorular, ocler, eslestirmeler, ogrenciler, anahtar) -> bytes:
    wb = Workbook()
    BLUE = "1A3A6B"; GREEN = "1E6B3A"; GREEN_L = "E8F5ED"
    RED  = "C0392B"; GRAY  = "F5F5F5"; WHITE   = "FFFFFF"

    def hstyle(cell):
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def borders(ws, r1, r2, c1, c2):
        t = Side(style="thin", color="CCCCCC")
        for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
            for cell in row:
                cell.border = Border(left=t, right=t, top=t, bottom=t)

    ws1 = wb.active
    ws1.title = "Question-LO Mapping"
    ws1.row_dimensions[1].height = 32
    for col, h in enumerate(["Q NO","QUESTION","ANSWER KEY","LO NO","LO DEFINITION","DIFFICULTY","SUCCESS %"], 1):
        hstyle(ws1.cell(row=1, column=col, value=h))

    oc_map = {o["no"]: o["tanim"] for o in ocler}
    for ri, s in enumerate(sorular, 2):
        esl    = eslestirmeler.get(s["no"], {})
        oc_no  = esl.get("oc_no", "")
        zorluk = esl.get("zorluk", "Medium")
        dogru  = anahtar[s["no"]-1] if anahtar and s["no"]-1 < len(anahtar) else "-"
        basari = None
        if ogrenciler and anahtar and s["no"]-1 < len(anahtar):
            cnt = sum(1 for o in ogrenciler if len(o["cevaplar"]) > s["no"]-1 and o["cevaplar"][s["no"]-1] == dogru)
            basari = cnt / len(ogrenciler)

        ws1.cell(ri,1,s["no"]).alignment = Alignment(horizontal="center")
        ws1.cell(ri,2,s["text"]).alignment = Alignment(wrap_text=True, vertical="top")
        ws1.cell(ri,3,dogru).alignment = Alignment(horizontal="center")
        ws1.cell(ri,4,oc_no).alignment = Alignment(horizontal="center")
        ws1.cell(ri,5,oc_map.get(oc_no,"")).alignment = Alignment(wrap_text=True)
        ws1.cell(ri,6,zorluk).alignment = Alignment(horizontal="center")
        pc = ws1.cell(ri,7,basari)
        pc.alignment = Alignment(horizontal="center")
        if basari is not None:
            pc.number_format = "0%"
            pc.fill = PatternFill("solid", fgColor=GREEN_L if basari>=0.6 else "FEF2F2")
            pc.font = Font(color=GREEN if basari>=0.6 else RED, bold=True)
        if ri % 2 == 0:
            for c in range(1,7):
                ws1.cell(ri,c).fill = PatternFill("solid", fgColor=GRAY)

    for col, w in zip("ABCDEFG",[9,58,13,10,44,12,12]):
        ws1.column_dimensions[col].width = w
    ws1.freeze_panes = "A2"
    borders(ws1, 1, len(sorular)+1, 1, 7)

    ws2 = wb.create_sheet("LO Summary")
    ws2.row_dimensions[1].height = 32
    for col, h in enumerate(["LO NO","LO DEFINITION","# QUESTIONS","AVG SUCCESS %","EVALUATION"],1):
        hstyle(ws2.cell(row=1, column=col, value=h))

    for ri, oc in enumerate(ocler, 2):
        oc_s = [s for s in sorular if eslestirmeler.get(s["no"],{}).get("oc_no")==oc["no"]]
        basarilar = []
        if ogrenciler and anahtar:
            for s in oc_s:
                dogru = anahtar[s["no"]-1] if s["no"]-1 < len(anahtar) else None
                if dogru:
                    cnt = sum(1 for o in ogrenciler if len(o["cevaplar"])>s["no"]-1 and o["cevaplar"][s["no"]-1]==dogru)
                    basarilar.append(cnt/len(ogrenciler))
        ort = sum(basarilar)/len(basarilar) if basarilar else None
        durum = ("Sufficient" if ort>=0.6 else "Needs Improvement") if ort is not None else ""

        ws2.cell(ri,1,oc["no"]).alignment = Alignment(horizontal="center")
        ws2.cell(ri,2,oc["tanim"]).alignment = Alignment(wrap_text=True)
        ws2.cell(ri,3,len(oc_s)).alignment = Alignment(horizontal="center")
        pc2 = ws2.cell(ri,4,ort)
        pc2.alignment = Alignment(horizontal="center")
        if ort is not None:
            pc2.number_format = "0%"
            pc2.fill = PatternFill("solid", fgColor=GREEN_L if ort>=0.6 else "FEF2F2")
            pc2.font = Font(color=GREEN if ort>=0.6 else RED, bold=True)
        ws2.cell(ri,5,durum).alignment = Alignment(horizontal="center")
        if ri % 2 == 0:
            for c in [1,2,3,5]:
                ws2.cell(ri,c).fill = PatternFill("solid", fgColor=GRAY)

    for col, w in zip("ABCDE",[10,52,14,16,20]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"
    borders(ws2, 1, len(ocler)+1, 1, 5)

    if ogrenciler:
        ws3 = wb.create_sheet("Student Results")
        ws3.row_dimensions[1].height = 32
        for col, h in enumerate(["NAME","STUDENT ID","ANSWERS","CORRECT","SCORE %"],1):
            hstyle(ws3.cell(row=1, column=col, value=h))
        for ri, o in enumerate(ogrenciler, 2):
            dogru_say = sum(1 for i,d in enumerate(anahtar) if i < len(o["cevaplar"]) and o["cevaplar"][i]==d) if anahtar else 0
            pct3 = dogru_say/len(anahtar) if anahtar else None
            ws3.cell(ri,1,o["ad"])
            ws3.cell(ri,2,o["no"]).alignment = Alignment(horizontal="center")
            ws3.cell(ri,3,o["cevaplar"]).alignment = Alignment(horizontal="center")
            ws3.cell(ri,4,dogru_say).alignment = Alignment(horizontal="center")
            p = ws3.cell(ri,5,pct3)
            p.alignment = Alignment(horizontal="center")
            if pct3 is not None: p.number_format = "0%"
            if ri % 2 == 0:
                for c in range(1,6): ws3.cell(ri,c).fill = PatternFill("solid",fgColor=GRAY)
        for col, w in zip("ABCDE",[26,14,40,14,12]):
            ws3.column_dimensions[col].width = w
        ws3.freeze_panes = "A2"
        borders(ws3,1,len(ogrenciler)+1,1,5)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Exam Outcome Mapper")
    st.markdown("---")
    steps = ["📂 Upload Exam","🎯 Define Outcomes","🔗 Map Questions","📊 Export Excel"]
    for i, s in enumerate(steps, 1):
        color = "#27ae60" if st.session_state.step > i else ("#1a3a6b" if st.session_state.step==i else "#aaa")
        icon  = "✅" if st.session_state.step > i else ("▶" if st.session_state.step==i else "○")
        st.markdown(f"<span style='color:{color};font-weight:600'>{icon} {s}</span>", unsafe_allow_html=True)
    st.markdown("---")
    if st.session_state.step > 1:
        if st.button("🔄 Start Over", use_container_width=True):
            for k in ["step","sorular","ocler","eslestirmeler","anahtar","ogrenciler"]:
                del st.session_state[k]
            st.rerun()
    st.markdown("---")
    st.markdown("<small style='color:#aaa'>Kapadokya University<br>Accreditation Report System<br><br>🤖 Powered by DeepSeek AI</small>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# STEP 1 — UPLOAD
# ══════════════════════════════════════════════════════════════
if st.session_state.step == 1:
    st.markdown("""<div class='step-header'>
        <h2>📂 Step 1 — Upload Exam File</h2>
        <p>System extracts questions automatically — each question must end with "?"</p>
    </div>""", unsafe_allow_html=True)

    uploaded = st.file_uploader("Choose file", type=["txt","pdf","docx","doc"])

    if uploaded:
        ext = uploaded.name.split(".")[-1].lower()
        with st.spinner("Extracting questions..."):
            try:
                sorular = []
                raw_bytes = uploaded.read()

                if ext == "txt":
                    text = raw_bytes.decode("utf-8", errors="ignore")
                    sorular = parse_text(text)
                    if not sorular:
                        sorular = parse_via_deepseek(text)
                elif ext in ["docx","doc"]:
                    from docx import Document
                    doc = Document(io.BytesIO(raw_bytes))
                    text = "\n".join(p.text for p in doc.paragraphs)
                    sorular = parse_text(text)
                    if not sorular:
                        sorular = parse_via_deepseek(text)
                elif ext == "pdf":
                    import pypdf
                    reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
                    text = "\n".join(page.extract_text() or "" for page in reader.pages)
                    sorular = parse_text(text)
                    if not sorular:
                        sorular = parse_via_deepseek(text)

                if not sorular:
                    st.error("❌ No questions found. Questions must end with '?'")
                else:
                    st.session_state.sorular = sorular
                    st.success(f"✅ **{len(sorular)} questions** extracted successfully!")
                    with st.expander(f"Preview ({len(sorular)} questions)", expanded=False):
                        for s in sorular:
                            st.markdown(f"**{s['no']}.** {s['text']}")
                    if st.button("Continue → Define Outcomes", type="primary", use_container_width=True):
                        st.session_state.step = 2
                        st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

# ══════════════════════════════════════════════════════════════
# STEP 2 — OUTCOMES
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 2:
    st.markdown("""<div class='step-header'>
        <h2>🎯 Step 2 — Define Learning Outcomes</h2>
        <p>Enter manually or upload a photo of your outcome list</p>
    </div>""", unsafe_allow_html=True)

    if st.session_state.ocler:
        st.markdown("**Defined Outcomes:**")
        for i, oc in enumerate(st.session_state.ocler):
            c1,c2,c3 = st.columns([1.5,7,1])
            with c1: st.markdown(f"<span class='oc-badge'>{oc['no']}</span>", unsafe_allow_html=True)
            with c2: st.caption(oc["tanim"])
            with c3:
                if st.button("×", key=f"del_{i}"):
                    st.session_state.ocler.pop(i); st.rerun()
        st.markdown("---")

    st.markdown("**Add New Outcome:**")
    c1,c2,c3 = st.columns([2,7,1.5])
    with c1:
        oc_no = st.text_input("No", value=f"LO-{len(st.session_state.ocler)+1}", label_visibility="collapsed")
    with c2:
        oc_tanim = st.text_input("Definition", label_visibility="collapsed", placeholder="Learning outcome definition...")
    with c3:
        if st.button("➕ Add", use_container_width=True):
            if oc_no and oc_tanim:
                if not any(o["no"]==oc_no for o in st.session_state.ocler):
                    st.session_state.ocler.append({"no":oc_no,"tanim":oc_tanim}); st.rerun()
                else: st.warning("This outcome number already exists.")
            else: st.warning("Please fill in both fields.")

    st.markdown("---")
    st.info("💡 **Tip:** You can also paste multiple outcomes at once below (one per line: `LO-1: definition`)"  )
    bulk = st.text_area("Bulk add outcomes (LO-1: definition, one per line)", height=100, label_visibility="collapsed", placeholder="LO-1: Explains machine learning algorithms\nLO-2: Applies supervised learning methods\nLO-3: Evaluates model performance")
    if st.button("➕ Add All", use_container_width=True):
        added = 0
        for line in bulk.strip().split("\n"):
            line = line.strip()
            if ":" in line:
                parts = line.split(":", 1)
                no = parts[0].strip()
                tanim = parts[1].strip()
                if no and tanim and not any(o["no"]==no for o in st.session_state.ocler):
                    st.session_state.ocler.append({"no": no, "tanim": tanim})
                    added += 1
        if added:
            st.success(f"✅ {added} outcomes added!")
            st.rerun()

    st.markdown("---")
    if st.session_state.ocler:
        if st.button("Continue → Map Questions", type="primary", use_container_width=True):
            st.session_state.step = 3; st.rerun()
    else:
        st.info("Add at least one learning outcome to continue.")

# ══════════════════════════════════════════════════════════════
# STEP 3 — MAP
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 3:
    st.markdown("""<div class='step-header'>
        <h2>🔗 Step 3 — Map Questions to Outcomes</h2>
        <p>Auto-map with AI or assign manually</p>
    </div>""", unsafe_allow_html=True)

    # ── OTOMATİK EŞLEŞTİRME BUTONU ──
    st.markdown("""<div class='auto-banner'>
        🤖 <strong>AI Auto-Mapping:</strong> DeepSeek reads all questions and learning outcomes,
        then automatically assigns the best match for each question.
    </div>""", unsafe_allow_html=True)

    if st.button("⚡ Auto-Map All Questions with AI", type="primary", use_container_width=True):
        chunks = (len(st.session_state.sorular) + 19) // 20
        with st.spinner(f"DeepSeek mapping {len(st.session_state.sorular)} questions in {chunks} batches — please wait..."):
            try:
                result = auto_match(st.session_state.sorular, st.session_state.ocler)
                st.session_state.eslestirmeler = result
                eslesen = sum(1 for v in result.values() if v.get("oc_no"))
                st.success(f"✅ {eslesen}/{len(st.session_state.sorular)} questions mapped automatically! Review below and adjust if needed.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Auto-mapping failed: {str(e)}")

    st.markdown("---")

    # Cevap anahtarı
    with st.expander("📝 Answer Key & Student Data (optional)", expanded=False):
        st.session_state.anahtar = st.text_input(
            "Answer Key (ABCDE... format)",
            value=st.session_state.anahtar, placeholder="ABCDEABCDE..."
        ).upper()
        ogr_raw = st.text_area("Student Answers (NAME \\t ID \\t ANSWERS per line)", height=120)
        if ogr_raw:
            ogrenciler = []
            for line in ogr_raw.strip().split("\n"):
                parts = re.split(r'\s{2,}|\t', line.strip())
                if len(parts) >= 2:
                    ogrenciler.append({"ad":" ".join(parts[:-2]),"no":parts[-2],"cevaplar":parts[-1].upper()})
            st.session_state.ogrenciler = ogrenciler
            st.caption(f"✅ {len(ogrenciler)} students loaded")

    # Soru listesi
    oc_keys   = [""] + [o["no"] for o in st.session_state.ocler]
    oc_labels = ["— Select LO —"] + [f"{o['no']} — {o['tanim']}" for o in st.session_state.ocler]
    zorluklar = ["Easy","Medium","Hard"]

    eslesen = sum(1 for v in st.session_state.eslestirmeler.values() if v.get("oc_no"))
    st.markdown(f"**{len(st.session_state.sorular)} Questions** — review and adjust if needed:")

    for s in st.session_state.sorular:
        esl = st.session_state.eslestirmeler.get(s["no"], {})
        matched = bool(esl.get("oc_no"))
        card_class = "soru-card matched" if matched else "soru-card"
        st.markdown(f"""<div class='{card_class}'>
            <div class='soru-no'>{'✅ ' if matched else ''}Question {s['no']}</div>
            <div class='soru-text'>{s['text']}</div>
        </div>""", unsafe_allow_html=True)

        c1,c2 = st.columns([3,1.5])
        with c1:
            cur = esl.get("oc_no","")
            idx = oc_keys.index(cur) if cur in oc_keys else 0
            sel = st.selectbox("lo", oc_keys, format_func=lambda x: oc_labels[oc_keys.index(x)],
                               index=idx, label_visibility="collapsed", key=f"oc_{s['no']}")
        with c2:
            cur_z = esl.get("zorluk","Medium")
            zi = zorluklar.index(cur_z) if cur_z in zorluklar else 1
            sel_z = st.selectbox("diff", zorluklar, index=zi, label_visibility="collapsed", key=f"z_{s['no']}")
        st.session_state.eslestirmeler[s["no"]] = {"oc_no": sel, "zorluk": sel_z}

    st.markdown("---")
    pct = eslesen/len(st.session_state.sorular) if st.session_state.sorular else 0
    st.progress(pct, text=f"Mapped: {eslesen}/{len(st.session_state.sorular)} questions")

    if st.button("📊 Generate Excel Report", type="primary", use_container_width=True):
        st.session_state.step = 4; st.rerun()

# ══════════════════════════════════════════════════════════════
# STEP 4 — EXCEL
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 4:
    st.markdown("""<div class='step-header'>
        <h2>✅ Step 4 — Report Ready</h2>
        <p>Your accreditation Excel report has been generated</p>
    </div>""", unsafe_allow_html=True)

    eslesen = sum(1 for v in st.session_state.eslestirmeler.values() if v.get("oc_no"))
    c1,c2,c3 = st.columns(3)
    with c1: st.metric("Total Questions", len(st.session_state.sorular))
    with c2: st.metric("Total Outcomes", len(st.session_state.ocler))
    with c3: st.metric("Mapped", f"{eslesen}/{len(st.session_state.sorular)}")

    with st.spinner("Building Excel..."):
        excel_bytes = build_excel(
            st.session_state.sorular, st.session_state.ocler,
            st.session_state.eslestirmeler, st.session_state.ogrenciler,
            st.session_state.anahtar
        )

    st.success("✅ 3-sheet Excel report is ready!")
    st.download_button(
        label="⬇ Download Excel Report (.xlsx)",
        data=excel_bytes,
        file_name="Exam-Outcome-Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary"
    )

    with st.expander("📊 Outcome Summary", expanded=True):
        for oc in st.session_state.ocler:
            n = sum(1 for s in st.session_state.sorular
                    if st.session_state.eslestirmeler.get(s["no"],{}).get("oc_no")==oc["no"])
            st.markdown(f"**{oc['no']}** — {oc['tanim']}")
            st.caption(f"{n} questions mapped")
            st.markdown("---")
