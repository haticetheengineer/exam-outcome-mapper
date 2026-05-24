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
.oc-badge {
    background: #1a3a6b; color: white; padding: 2px 10px;
    border-radius: 12px; font-size: 0.72rem; font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

def init_state():
    defaults = {
        "step": 1, "sorular": [], "ocler": [],
        "eslestirmeler": {}, "anahtar": "", "ogrenciler": [],
        "auto_mapped": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

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

def auto_match(sorular, ocler):
    oc_listesi = "\n".join([f"- {o['no']}: {o['tanim']}" for o in ocler])
    result = {}
    chunk_size = 20

    for i in range(0, len(sorular), chunk_size):
        chunk = sorular[i:i+chunk_size]
        soru_listesi = "\n".join([f"{s['no']}. {s['text']}" for s in chunk])

        prompt = f"""Match each exam question to the most suitable learning outcome.

LEARNING OUTCOMES:
{oc_listesi}

QUESTIONS:
{soru_listesi}

Return ONLY valid JSON, nothing else:
{{"eslestirmeler":[{{"soru_no":1,"oc_no":"LO-1","zorluk":"Medium"}}]}}"""

        raw = deepseek_chat([{"role": "user", "content": prompt}], max_tokens=2048)
        raw = re.sub(r'```json|```', '', raw).strip()
        # JSON bulmaya calis
        match = re.search(r'\{.*\}', raw, re.DOTALL)
        if match:
            data = json.loads(match.group())
            for item in data.get("eslestirmeler", []):
                try:
                    soru_no = int(item["soru_no"])
                    result[soru_no] = {
                        "oc_no": str(item.get("oc_no", "")),
                        "zorluk": str(item.get("zorluk", "Medium"))
                    }
                except:
                    pass
    return result

def parse_text(text):
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

def parse_via_deepseek(text):
    prompt = f"""Extract all questions ending with "?" from this exam text.
Return ONLY JSON:
{{"sorular":[{{"no":1,"text":"Question?"}}]}}

TEXT:
{text[:10000]}"""
    raw = deepseek_chat([{"role": "user", "content": prompt}])
    raw = re.sub(r'```json|```', '', raw).strip()
    return json.loads(raw).get("sorular", [])

def build_excel(sorular, ocler, eslestirmeler, ogrenciler, anahtar):
    wb = Workbook()
    BLUE="1A3A6B"; GREEN="1E6B3A"; GREEN_L="E8F5ED"
    RED="C0392B"; GRAY="F5F5F5"; WHITE="FFFFFF"

    def hstyle(cell):
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def borders(ws, r1, r2, c1, c2):
        t = Side(style="thin", color="CCCCCC")
        for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
            for cell in row:
                cell.border = Border(left=t, right=t, top=t, bottom=t)

    def get_esl(no):
        return eslestirmeler.get(no, eslestirmeler.get(str(no), {}))

    ws1 = wb.active
    ws1.title = "Question-LO Mapping"
    ws1.row_dimensions[1].height = 32
    for col, h in enumerate(["Q NO","QUESTION","ANSWER KEY","LO NO","LO DEFINITION","DIFFICULTY","SUCCESS %"],1):
        hstyle(ws1.cell(row=1, column=col, value=h))

    oc_map = {o["no"]: o["tanim"] for o in ocler}
    for ri, s in enumerate(sorular, 2):
        esl = get_esl(s["no"])
        oc_no = esl.get("oc_no","")
        zorluk = esl.get("zorluk","Medium")
        dogru = anahtar[s["no"]-1] if anahtar and s["no"]-1 < len(anahtar) else "-"
        basari = None
        if ogrenciler and anahtar and s["no"]-1 < len(anahtar):
            cnt = sum(1 for o in ogrenciler if len(o["cevaplar"])>s["no"]-1 and o["cevaplar"][s["no"]-1]==dogru)
            basari = cnt/len(ogrenciler)

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
        if ri%2==0:
            for c in range(1,7): ws1.cell(ri,c).fill = PatternFill("solid",fgColor=GRAY)

    for col,w in zip("ABCDEFG",[9,58,13,10,44,12,12]):
        ws1.column_dimensions[col].width = w
    ws1.freeze_panes = "A2"
    borders(ws1,1,len(sorular)+1,1,7)

    ws2 = wb.create_sheet("LO Summary")
    ws2.row_dimensions[1].height = 32
    for col,h in enumerate(["LO NO","LO DEFINITION","# QUESTIONS","AVG SUCCESS %","EVALUATION"],1):
        hstyle(ws2.cell(row=1,column=col,value=h))

    for ri, oc in enumerate(ocler,2):
        oc_s = [s for s in sorular if get_esl(s["no"]).get("oc_no")==oc["no"]]
        basarilar = []
        if ogrenciler and anahtar:
            for s in oc_s:
                dogru = anahtar[s["no"]-1] if s["no"]-1<len(anahtar) else None
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
            pc2.fill = PatternFill("solid",fgColor=GREEN_L if ort>=0.6 else "FEF2F2")
            pc2.font = Font(color=GREEN if ort>=0.6 else RED, bold=True)
        ws2.cell(ri,5,durum).alignment = Alignment(horizontal="center")
        if ri%2==0:
            for c in [1,2,3,5]: ws2.cell(ri,c).fill = PatternFill("solid",fgColor=GRAY)

    for col,w in zip("ABCDE",[10,52,14,16,20]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"
    borders(ws2,1,len(ocler)+1,1,5)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Exam Outcome Mapper")
    st.markdown("---")
    steps = ["📂 Upload","🎯 Outcomes","🔗 Map","📊 Excel"]
    for i, s in enumerate(steps,1):
        color = "#27ae60" if st.session_state.step>i else ("#1a3a6b" if st.session_state.step==i else "#aaa")
        icon = "✅" if st.session_state.step>i else ("▶" if st.session_state.step==i else "○")
        st.markdown(f"<span style='color:{color};font-weight:600'>{icon} {s}</span>", unsafe_allow_html=True)
    st.markdown("---")
    if st.session_state.step > 1:
        if st.button("🔄 Start Over", use_container_width=True):
            for k in ["step","sorular","ocler","eslestirmeler","anahtar","ogrenciler","auto_mapped"]:
                del st.session_state[k]
            st.rerun()
    st.markdown("---")
    st.markdown("<small style='color:#aaa'>Kapadokya University<br>🤖 DeepSeek AI</small>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# STEP 1
# ══════════════════════════════════════════════════════════════
if st.session_state.step == 1:
    st.markdown("""<div class='step-header'>
        <h2>📂 Step 1 — Upload Exam File</h2>
        <p>Questions must end with "?"</p>
    </div>""", unsafe_allow_html=True)

    uploaded = st.file_uploader("Choose file", type=["txt","pdf","docx","doc"])
    if uploaded:
        ext = uploaded.name.split(".")[-1].lower()
        with st.spinner("Extracting questions..."):
            try:
                raw_bytes = uploaded.read()
                sorular = []
                if ext == "txt":
                    text = raw_bytes.decode("utf-8", errors="ignore")
                    sorular = parse_text(text) or parse_via_deepseek(text)
                elif ext in ["docx","doc"]:
                    from docx import Document
                    text = "\n".join(p.text for p in Document(io.BytesIO(raw_bytes)).paragraphs)
                    sorular = parse_text(text) or parse_via_deepseek(text)
                elif ext == "pdf":
                    import pypdf
                    text = "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(io.BytesIO(raw_bytes)).pages)
                    sorular = parse_text(text) or parse_via_deepseek(text)

                if not sorular:
                    st.error("❌ No questions found. Questions must end with '?'")
                else:
                    st.session_state.sorular = sorular
                    st.success(f"✅ {len(sorular)} questions extracted!")
                    with st.expander("Preview", expanded=False):
                        for s in sorular:
                            st.markdown(f"**{s['no']}.** {s['text']}")
                    if st.button("Continue →", type="primary", use_container_width=True):
                        st.session_state.step = 2
                        st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

# ══════════════════════════════════════════════════════════════
# STEP 2
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 2:
    st.markdown("""<div class='step-header'>
        <h2>🎯 Step 2 — Define Learning Outcomes</h2>
        <p>Enter manually or paste in bulk</p>
    </div>""", unsafe_allow_html=True)

    if st.session_state.ocler:
        st.markdown("**Current Outcomes:**")
        for i, oc in enumerate(st.session_state.ocler):
            c1,c2,c3 = st.columns([1.5,7,1])
            with c1: st.markdown(f"<span class='oc-badge'>{oc['no']}</span>", unsafe_allow_html=True)
            with c2: st.caption(oc["tanim"])
            with c3:
                if st.button("×", key=f"del_{i}"):
                    st.session_state.ocler.pop(i); st.rerun()
        st.markdown("---")

    c1,c2,c3 = st.columns([2,7,1.5])
    with c1:
        oc_no = st.text_input("No", value=f"LO-{len(st.session_state.ocler)+1}", label_visibility="collapsed")
    with c2:
        oc_tanim = st.text_input("Def", label_visibility="collapsed", placeholder="Learning outcome...")
    with c3:
        if st.button("➕", use_container_width=True):
            if oc_no and oc_tanim and not any(o["no"]==oc_no for o in st.session_state.ocler):
                st.session_state.ocler.append({"no":oc_no,"tanim":oc_tanim}); st.rerun()

    st.markdown("---")
    st.info("💡 Paste multiple outcomes (one per line: `LO-1: definition`)")
    bulk = st.text_area("Bulk", height=120, label_visibility="collapsed",
                        placeholder="LO-1: Explains ML algorithms\nLO-2: Applies methods")
    if st.button("➕ Add All", use_container_width=True):
        added = 0
        for line in bulk.strip().split("\n"):
            if ":" in line:
                no, tanim = line.split(":",1)
                no, tanim = no.strip(), tanim.strip()
                if no and tanim and not any(o["no"]==no for o in st.session_state.ocler):
                    st.session_state.ocler.append({"no":no,"tanim":tanim}); added+=1
        if added:
            st.success(f"✅ {added} added!"); st.rerun()

    st.markdown("---")
    if st.session_state.ocler:
        if st.button("Continue →", type="primary", use_container_width=True):
            st.session_state.step = 3; st.rerun()
    else:
        st.info("Add at least one outcome.")

# ══════════════════════════════════════════════════════════════
# STEP 3
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 3:
    st.markdown("""<div class='step-header'>
        <h2>🔗 Step 3 — Map Questions to Outcomes</h2>
        <p>Use AI auto-map then review</p>
    </div>""", unsafe_allow_html=True)

    # AUTO MAP
    if st.button("⚡ Auto-Map All with AI", type="primary", use_container_width=True):
        n = len(st.session_state.sorular)
        chunks = (n + 19) // 20
        progress = st.progress(0, text="Starting...")
        errors = []

        new_map = {}
        for i in range(0, n, 20):
            chunk = st.session_state.sorular[i:i+20]
            batch_num = i//20 + 1
            progress.progress(batch_num/chunks, text=f"Batch {batch_num}/{chunks}...")
            try:
                oc_listesi = "\n".join([f"- {o['no']}: {o['tanim']}" for o in st.session_state.ocler])
                soru_listesi = "\n".join([f"{s['no']}. {s['text']}" for s in chunk])
                prompt = f"""Match each question to the best learning outcome.

LEARNING OUTCOMES:
{oc_listesi}

QUESTIONS:
{soru_listesi}

Return ONLY this JSON format:
{{"eslestirmeler":[{{"soru_no":1,"oc_no":"LO-1","zorluk":"Medium"}}]}}"""

                raw = deepseek_chat([{"role":"user","content":prompt}], max_tokens=2048)
                raw = re.sub(r'```json|```','',raw).strip()
                m = re.search(r'\{.*\}', raw, re.DOTALL)
                if m:
                    data = json.loads(m.group())
                    for item in data.get("eslestirmeler",[]):
                        soru_no = int(item["soru_no"])
                        new_map[soru_no] = {
                            "oc_no": str(item.get("oc_no","")),
                            "zorluk": str(item.get("zorluk","Medium"))
                        }
            except Exception as e:
                errors.append(f"Batch {batch_num}: {e}")

        progress.empty()

        # Session state'e yaz
        st.session_state.eslestirmeler = new_map
        st.session_state.auto_mapped = True

        mapped = sum(1 for v in new_map.values() if v.get("oc_no"))
        if errors:
            st.warning(f"⚠️ {mapped} mapped, {len(errors)} errors: {errors[0]}")
        else:
            st.success(f"✅ {mapped}/{n} questions mapped!")

    st.markdown("---")

    # MEVCUT EŞLEŞTİRMELERİ GÖSTER
    oc_keys   = [""] + [o["no"] for o in st.session_state.ocler]
    oc_labels = ["— Select —"] + [f"{o['no']}: {o['tanim'][:40]}..." if len(o['tanim'])>40 else f"{o['no']}: {o['tanim']}" for o in st.session_state.ocler]
    zorluklar = ["Easy","Medium","Hard"]

    mapped_count = sum(1 for v in st.session_state.eslestirmeler.values() if v.get("oc_no"))
    st.progress(mapped_count/len(st.session_state.sorular) if st.session_state.sorular else 0,
                text=f"Mapped: {mapped_count}/{len(st.session_state.sorular)}")

    # Her soru için widget — key olarak soru no kullan
    for s in st.session_state.sorular:
        sno = s["no"]
        esl = st.session_state.eslestirmeler.get(sno, {})
        cur_oc = esl.get("oc_no","")
        cur_zor = esl.get("zorluk","Medium")

        matched = bool(cur_oc)
        prefix = "✅" if matched else "⬜"

        with st.container():
            st.markdown(f"**{prefix} Q{sno}.** {s['text'][:100]}{'...' if len(s['text'])>100 else ''}")
            c1,c2 = st.columns([3,1.5])
            with c1:
                idx = oc_keys.index(cur_oc) if cur_oc in oc_keys else 0
                sel = st.selectbox(
                    f"LO_{sno}",
                    options=oc_keys,
                    format_func=lambda x: oc_labels[oc_keys.index(x)],
                    index=idx,
                    label_visibility="collapsed",
                    key=f"sel_oc_{sno}"
                )
            with c2:
                zi = zorluklar.index(cur_zor) if cur_zor in zorluklar else 1
                sel_z = st.selectbox(
                    f"Z_{sno}",
                    options=zorluklar,
                    index=zi,
                    label_visibility="collapsed",
                    key=f"sel_z_{sno}"
                )

            # Her zaman session state'e yaz (widget değeri kazanır)
            st.session_state.eslestirmeler[sno] = {"oc_no": sel, "zorluk": sel_z}

    st.markdown("---")
    with st.expander("📝 Answer Key & Students (optional)"):
        st.session_state.anahtar = st.text_input("Answer Key", value=st.session_state.anahtar, placeholder="ABCDE...").upper()
        ogr_raw = st.text_area("Students (NAME\\tID\\tANSWERS)", height=100)
        if ogr_raw:
            ogrenciler = []
            for line in ogr_raw.strip().split("\n"):
                parts = re.split(r'\s{2,}|\t', line.strip())
                if len(parts) >= 2:
                    ogrenciler.append({"ad":" ".join(parts[:-2]),"no":parts[-2],"cevaplar":parts[-1].upper()})
            st.session_state.ogrenciler = ogrenciler

    if st.button("📊 Generate Excel", type="primary", use_container_width=True):
        st.session_state.step = 4; st.rerun()

# ══════════════════════════════════════════════════════════════
# STEP 4
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 4:
    st.markdown("""<div class='step-header'>
        <h2>✅ Report Ready</h2>
        <p>Download your accreditation Excel report</p>
    </div>""", unsafe_allow_html=True)

    mapped = sum(1 for v in st.session_state.eslestirmeler.values() if v.get("oc_no"))
    c1,c2,c3 = st.columns(3)
    with c1: st.metric("Questions", len(st.session_state.sorular))
    with c2: st.metric("Outcomes", len(st.session_state.ocler))
    with c3: st.metric("Mapped", f"{mapped}/{len(st.session_state.sorular)}")

    with st.spinner("Building Excel..."):
        excel_bytes = build_excel(
            st.session_state.sorular, st.session_state.ocler,
            st.session_state.eslestirmeler, st.session_state.ogrenciler,
            st.session_state.anahtar
        )

    st.success("✅ Excel ready!")
    st.download_button(
        "⬇ Download Excel (.xlsx)", excel_bytes,
        "Exam-Outcome-Report.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary"
    )

    with st.expander("📊 LO Summary", expanded=True):
        for oc in st.session_state.ocler:
            n = sum(1 for s in st.session_state.sorular
                    if st.session_state.eslestirmeler.get(s["no"],{}).get("oc_no")==oc["no"])
            st.markdown(f"**{oc['no']}** — {oc['tanim']}: **{n} questions**")
