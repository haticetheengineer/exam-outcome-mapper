import streamlit as st
import json
import re
import io
import base64
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Sınav ÖÇ Eşleştirme", page_icon="📋", layout="centered")

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
.soru-no { color: #c0392b; font-weight: 700; font-size: 0.85rem; }
.soru-text { font-size: 0.88rem; line-height: 1.5; margin-top: 4px; }
.oc-badge {
    background: #1a3a6b; color: white; padding: 2px 10px;
    border-radius: 12px; font-size: 0.72rem; font-weight: 600;
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

# ── GPT API ───────────────────────────────────────────────────
@st.cache_resource
def get_client():
    return OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

def gpt_chat(messages: list, max_tokens=4096) -> str:
    client = get_client()
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=messages,
        max_tokens=max_tokens
    )
    return resp.choices[0].message.content.strip()

def gpt_vision(img_bytes: bytes, mime: str, prompt: str) -> str:
    client = get_client()
    b64 = base64.b64encode(img_bytes).decode()
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": [
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
            {"type": "text", "text": prompt}
        ]}],
        max_tokens=2000
    )
    return resp.choices[0].message.content.strip()

def gpt_pdf(pdf_bytes: bytes, prompt: str) -> str:
    """PDF'i metin olarak çıkarıp GPT'ye gönder"""
    import pypdf
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    full_prompt = f"{prompt}\n\nMETİN:\n{text[:12000]}"
    return gpt_chat([{"role": "user", "content": full_prompt}])

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

def parse_via_gpt(text: str) -> list:
    prompt = f"""Aşağıdaki sınav metninden "?" ile biten TÜM soruları çıkar.
SADECE JSON döndür, başka hiçbir şey yazma:
{{"sorular":[{{"no":1,"text":"Soru metni?"}},{{"no":2,"text":"..."}}]}}

METİN:
{text[:10000]}"""
    raw = gpt_chat([{"role": "user", "content": prompt}])
    raw = re.sub(r'```json|```', '', raw).strip()
    return json.loads(raw).get("sorular", [])

def read_oc_from_image(img_bytes: bytes, mime: str) -> list:
    prompt = """Bu resimde öğrenim çıktıları (ÖÇ) listesi var.
Tüm ÖÇ'leri çıkar. SADECE JSON döndür:
{"ocler":[{"no":"ÖÇ-1","tanim":"..."},{"no":"ÖÇ-2","tanim":"..."}]}"""
    raw = gpt_vision(img_bytes, mime, prompt)
    raw = re.sub(r'```json|```', '', raw).strip()
    return json.loads(raw).get("ocler", [])

def build_excel(sorular, ocler, eslestirmeler, ogrenciler, anahtar) -> bytes:
    wb = Workbook()
    BLUE  = "1A3A6B"; GREEN = "1E6B3A"; GREEN_L = "E8F5ED"
    RED   = "C0392B"; GRAY  = "F5F5F5"; WHITE   = "FFFFFF"

    def hstyle(cell):
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def borders(ws, r1, r2, c1, c2):
        t = Side(style="thin", color="CCCCCC")
        for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
            for cell in row:
                cell.border = Border(left=t, right=t, top=t, bottom=t)

    # SAYFA 1: Soru-ÖÇ
    ws1 = wb.active
    ws1.title = "Soru-ÖÇ Eşleştirme"
    ws1.row_dimensions[1].height = 32
    for col, h in enumerate(["SORU NO","SORU METNİ","DOĞRU CEVAP","ÖÇ NO","ÖÇ TANIMI","ZORLUK","BAŞARI %"], 1):
        hstyle(ws1.cell(row=1, column=col, value=h))

    oc_map = {o["no"]: o["tanim"] for o in ocler}
    for ri, s in enumerate(sorular, 2):
        esl    = eslestirmeler.get(s["no"], {})
        oc_no  = esl.get("oc_no", "")
        zorluk = esl.get("zorluk", "Orta")
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

    # SAYFA 2: ÖÇ Özet
    ws2 = wb.create_sheet("ÖÇ Özet")
    ws2.row_dimensions[1].height = 32
    for col, h in enumerate(["ÖÇ NO","ÖÇ TANIMI","SORU SAYISI","ORT. BAŞARI %","DEĞERLENDİRME"],1):
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
        durum = ("✅ Yeterli" if ort>=0.6 else "⚠️ Geliştirilmeli") if ort is not None else ""

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

    # SAYFA 3: Öğrenci
    if ogrenciler:
        ws3 = wb.create_sheet("Öğrenci Sonuçları")
        ws3.row_dimensions[1].height = 32
        for col, h in enumerate(["AD SOYAD","ÖĞRENCİ NO","CEVAPLAR","DOĞRU SAYISI","PUAN %"],1):
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
    st.markdown("### 📋 Sınav ÖÇ Sistemi")
    st.markdown("---")
    steps = ["📂 Sınav Yükle","🎯 ÖÇ Tanımla","🔗 Eşleştir","📊 Excel"]
    for i, s in enumerate(steps, 1):
        color = "#27ae60" if st.session_state.step > i else ("#1a3a6b" if st.session_state.step==i else "#aaa")
        icon  = "✅" if st.session_state.step > i else ("▶" if st.session_state.step==i else "○")
        st.markdown(f"<span style='color:{color};font-weight:600'>{icon} {s}</span>", unsafe_allow_html=True)
    st.markdown("---")
    if st.session_state.step > 1:
        if st.button("🔄 Başa Dön", use_container_width=True):
            for k in ["step","sorular","ocler","eslestirmeler","anahtar","ogrenciler"]:
                del st.session_state[k]
            st.rerun()
    st.markdown("---")
    st.markdown("<small style='color:#aaa'>Kapadokya Üniversitesi<br>Akreditasyon Raporu Sistemi<br><br>🤖 GPT-4o mini ile çalışır</small>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ADIM 1 — DOSYA YÜKLE
# ══════════════════════════════════════════════════════════════
if st.session_state.step == 1:
    st.markdown("""<div class='step-header'>
        <h2>📂 Adım 1 — Sınav Dosyasını Yükle</h2>
        <p>Sistem soruları otomatik çıkarır — her soru "?" ile bitmelidir</p>
    </div>""", unsafe_allow_html=True)

    uploaded = st.file_uploader("Dosya seç", type=["txt","pdf","docx","doc"])

    if uploaded:
        ext = uploaded.name.split(".")[-1].lower()
        with st.spinner("Sorular çıkarılıyor..."):
            try:
                sorular = []
                raw_bytes = uploaded.read()

                if ext == "txt":
                    text = raw_bytes.decode("utf-8", errors="ignore")
                    sorular = parse_text(text)
                    if not sorular:
                        sorular = parse_via_gpt(text)

                elif ext in ["docx","doc"]:
                    from docx import Document
                    doc = Document(io.BytesIO(raw_bytes))
                    text = "\n".join(p.text for p in doc.paragraphs)
                    sorular = parse_text(text)
                    if not sorular:
                        sorular = parse_via_gpt(text)

                elif ext == "pdf":
                    import pypdf
                    reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
                    text = "\n".join(page.extract_text() or "" for page in reader.pages)
                    sorular = parse_text(text)
                    if not sorular:
                        sorular = parse_via_gpt(text)

                if not sorular:
                    st.error("❌ Hiç soru bulunamadı. Sorular '?' ile bitmelidir.")
                else:
                    st.session_state.sorular = sorular
                    st.success(f"✅ **{len(sorular)} soru** başarıyla çıkarıldı!")
                    with st.expander(f"Bulunan sorular ({len(sorular)} adet)", expanded=False):
                        for s in sorular:
                            st.markdown(f"**{s['no']}.** {s['text']}")
                    if st.button("Devam → ÖÇ Tanımla", type="primary", use_container_width=True):
                        st.session_state.step = 2
                        st.rerun()
            except Exception as e:
                st.error(f"❌ Hata: {str(e)}")

# ══════════════════════════════════════════════════════════════
# ADIM 2 — ÖÇ TANIMLA
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 2:
    st.markdown("""<div class='step-header'>
        <h2>🎯 Adım 2 — Öğrenim Çıktılarını Tanımla</h2>
        <p>Elle girin veya ÖÇ listesinin fotoğrafını yükleyin</p>
    </div>""", unsafe_allow_html=True)

    if st.session_state.ocler:
        st.markdown("**Tanımlanan ÖÇ'ler:**")
        for i, oc in enumerate(st.session_state.ocler):
            c1,c2,c3 = st.columns([1.5,7,1])
            with c1: st.markdown(f"<span class='oc-badge'>{oc['no']}</span>", unsafe_allow_html=True)
            with c2: st.caption(oc["tanim"])
            with c3:
                if st.button("×", key=f"del_{i}"):
                    st.session_state.ocler.pop(i); st.rerun()
        st.markdown("---")

    st.markdown("**Yeni ÖÇ Ekle:**")
    c1,c2,c3 = st.columns([2,7,1.5])
    with c1:
        oc_no = st.text_input("No", value=f"ÖÇ-{len(st.session_state.ocler)+1}", label_visibility="collapsed")
    with c2:
        oc_tanim = st.text_input("Tanım", label_visibility="collapsed", placeholder="Öğrenim çıktısı tanımı...")
    with c3:
        if st.button("➕ Ekle", use_container_width=True):
            if oc_no and oc_tanim:
                if not any(o["no"]==oc_no for o in st.session_state.ocler):
                    st.session_state.ocler.append({"no":oc_no,"tanim":oc_tanim}); st.rerun()
                else: st.warning("Bu ÖÇ numarası zaten var.")
            else: st.warning("ÖÇ no ve tanımını doldurun.")

    st.markdown("---")
    st.markdown("**📷 ÖÇ listesinin fotoğrafından otomatik oku:**")
    oc_img = st.file_uploader("ÖÇ resmi", type=["jpg","jpeg","png","webp"], label_visibility="collapsed")
    if oc_img:
        with st.spinner("Resimden ÖÇ'ler okunuyor..."):
            try:
                new_ocler = read_oc_from_image(oc_img.read(), oc_img.type)
                added = 0
                for oc in new_ocler:
                    if not any(o["no"]==oc["no"] for o in st.session_state.ocler):
                        st.session_state.ocler.append(oc); added += 1
                st.success(f"✅ {added} yeni ÖÇ eklendi!"); st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

    st.markdown("---")
    if st.session_state.ocler:
        if st.button("Devam → Soruları Eşleştir", type="primary", use_container_width=True):
            st.session_state.step = 3; st.rerun()
    else:
        st.info("En az bir ÖÇ ekleyin, sonra devam edebilirsiniz.")

# ══════════════════════════════════════════════════════════════
# ADIM 3 — EŞLEŞTİR
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 3:
    st.markdown("""<div class='step-header'>
        <h2>🔗 Adım 3 — Soru ÖÇ Eşleştirme</h2>
        <p>Her soruya ÖÇ ve zorluk seviyesi atayın</p>
    </div>""", unsafe_allow_html=True)

    with st.expander("📝 Cevap Anahtarı ve Öğrenci Verileri (opsiyonel)", expanded=False):
        st.session_state.anahtar = st.text_input(
            "Cevap Anahtarı (ABCDE... formatında)",
            value=st.session_state.anahtar, placeholder="ABCDEABCDE..."
        ).upper()
        ogr_raw = st.text_area("Öğrenci Cevapları (AD SOYAD \\t NO \\t CEVAPLAR)", height=120)
        if ogr_raw:
            ogrenciler = []
            for line in ogr_raw.strip().split("\n"):
                parts = re.split(r'\s{2,}|\t', line.strip())
                if len(parts) >= 2:
                    ogrenciler.append({"ad":" ".join(parts[:-2]),"no":parts[-2],"cevaplar":parts[-1].upper()})
            st.session_state.ogrenciler = ogrenciler
            st.caption(f"✅ {len(ogrenciler)} öğrenci yüklendi")

    oc_keys   = [""] + [o["no"] for o in st.session_state.ocler]
    oc_labels = ["— ÖÇ Seç —"] + [f"{o['no']} — {o['tanim']}" for o in st.session_state.ocler]
    zorluklar = ["Kolay","Orta","Zor"]

    st.markdown(f"**{len(st.session_state.sorular)} Soru:**")
    for s in st.session_state.sorular:
        st.markdown(f"""<div class='soru-card'>
            <div class='soru-no'>Soru {s['no']}</div>
            <div class='soru-text'>{s['text']}</div>
        </div>""", unsafe_allow_html=True)
        esl = st.session_state.eslestirmeler.get(s["no"], {})
        c1,c2 = st.columns([3,1.5])
        with c1:
            cur = esl.get("oc_no","")
            idx = oc_keys.index(cur) if cur in oc_keys else 0
            sel = st.selectbox("oc", oc_keys, format_func=lambda x: oc_labels[oc_keys.index(x)],
                               index=idx, label_visibility="collapsed", key=f"oc_{s['no']}")
        with c2:
            cur_z = esl.get("zorluk","Orta")
            zi = zorluklar.index(cur_z) if cur_z in zorluklar else 1
            sel_z = st.selectbox("zor", zorluklar, index=zi, label_visibility="collapsed", key=f"z_{s['no']}")
        st.session_state.eslestirmeler[s["no"]] = {"oc_no": sel, "zorluk": sel_z}

    eslesen = sum(1 for v in st.session_state.eslestirmeler.values() if v.get("oc_no"))
    st.markdown("---")
    pct = eslesen/len(st.session_state.sorular) if st.session_state.sorular else 0
    st.progress(pct, text=f"Eşleştirme: {eslesen}/{len(st.session_state.sorular)} soru")

    if st.button("📊 Excel Raporu Oluştur", type="primary", use_container_width=True):
        st.session_state.step = 4; st.rerun()

# ══════════════════════════════════════════════════════════════
# ADIM 4 — EXCEL
# ══════════════════════════════════════════════════════════════
elif st.session_state.step == 4:
    st.markdown("""<div class='step-header'>
        <h2>✅ Adım 4 — Rapor Hazır</h2>
        <p>Excel dosyası oluşturuldu, indirmeye hazır</p>
    </div>""", unsafe_allow_html=True)

    eslesen = sum(1 for v in st.session_state.eslestirmeler.values() if v.get("oc_no"))
    c1,c2,c3 = st.columns(3)
    with c1: st.metric("Toplam Soru", len(st.session_state.sorular))
    with c2: st.metric("Toplam ÖÇ", len(st.session_state.ocler))
    with c3: st.metric("Eşleştirilen", f"{eslesen}/{len(st.session_state.sorular)}")

    with st.spinner("Excel hazırlanıyor..."):
        excel_bytes = build_excel(
            st.session_state.sorular, st.session_state.ocler,
            st.session_state.eslestirmeler, st.session_state.ogrenciler,
            st.session_state.anahtar
        )

    st.success("✅ 3 sayfalık Excel raporu hazır!")
    st.download_button(
        label="⬇ Excel Raporu İndir (.xlsx)",
        data=excel_bytes,
        file_name="Sinav-OC-Raporu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary"
    )

    with st.expander("📊 ÖÇ Özet", expanded=True):
        for oc in st.session_state.ocler:
            n = sum(1 for s in st.session_state.sorular
                    if st.session_state.eslestirmeler.get(s["no"],{}).get("oc_no")==oc["no"])
            st.markdown(f"**{oc['no']}** — {oc['tanim']}")
            st.caption(f"{n} soru eşleşti")
            st.markdown("---")
