"""
Research Co-Pilot (v5 — + Brainstorming Judul)
----------------------------------------------
0. Brainstorming Judul  : dari topik -> beberapa alternatif judul yang bisa dipilih.
1. Cari referensi NYATA  : OpenAlex + Semantic Scholar, verifikasi DOI (Crossref).
2. Draf BAB I / II / III : mengalir, dari referensi terverifikasi, sitasi asli.
3. Unduh Word (.docx) + Daftar Pustaka otomatis.

Prinsip: sitasi nyata (bukan karangan), TIDAK mengarang data/hasil.
BAB IV (Hasil) = data penelitian Anda sendiri, tidak dibuat otomatis.

Env var Railway: ANTHROPIC_API_KEY (wajib), SEMANTIC_SCHOLAR_API_KEY (opsional)
requirements.txt: streamlit, requests, anthropic, python-docx
"""

import os
import io
import requests
import streamlit as st

MAILTO = "bayushadega74@gmail.com"
OPENALEX_KEY = None

st.set_page_config(page_title="Research Co-Pilot", page_icon="📚", layout="centered")


# ============ SUMBER + VERIFIKASI ============
def reconstruct_abstract(inv):
    if not inv:
        return ""
    pos = {}
    for word, idxs in inv.items():
        for i in idxs:
            pos[i] = word
    return " ".join(pos[i] for i in sorted(pos))


@st.cache_data(show_spinner=False)
def search_openalex(query, per_page=20):
    params = {"search": query, "per-page": per_page, "mailto": MAILTO}
    if OPENALEX_KEY:
        params["api_key"] = OPENALEX_KEY
    try:
        r = requests.get("https://api.openalex.org/works", params=params, timeout=25)
        if r.status_code != 200:
            return []
        out = []
        for w in r.json().get("results", []):
            doi = (w.get("doi") or "").replace("https://doi.org/", "").lower()
            src = (w.get("primary_location") or {}).get("source") or {}
            out.append({
                "title": w.get("title") or "(tanpa judul)",
                "year": w.get("publication_year"),
                "doi": doi,
                "authors": [a["author"]["display_name"]
                            for a in w.get("authorships", [])][:5],
                "venue": src.get("display_name") or "",
                "cited_by": w.get("cited_by_count", 0) or 0,
                "abstract": reconstruct_abstract(w.get("abstract_inverted_index")),
            })
        return out
    except requests.RequestException:
        return []


@st.cache_data(show_spinner=False)
def search_semantic_scholar(query, limit=20):
    url = "https://api.semanticscholar.org/graph/v1/paper/search"
    fields = "title,year,abstract,venue,citationCount,externalIds,authors"
    params = {"query": query, "limit": limit, "fields": fields}
    headers = {}
    key = os.environ.get("SEMANTIC_SCHOLAR_API_KEY")
    if key:
        headers["x-api-key"] = key
    try:
        r = requests.get(url, params=params, headers=headers, timeout=25)
        if r.status_code != 200:
            return []
        out = []
        for w in r.json().get("data", []):
            doi = ((w.get("externalIds") or {}).get("DOI") or "").lower()
            out.append({
                "title": w.get("title") or "(tanpa judul)",
                "year": w.get("year"),
                "doi": doi,
                "authors": [a.get("name") for a in (w.get("authors") or [])][:5],
                "venue": w.get("venue") or "",
                "cited_by": w.get("citationCount", 0) or 0,
                "abstract": w.get("abstract") or "",
            })
        return out
    except requests.RequestException:
        return []


def merge_dedupe(*lists):
    by_doi = {}
    for lst in lists:
        for r in lst:
            doi = (r.get("doi") or "").strip().lower()
            if not doi:
                continue
            if doi not in by_doi:
                by_doi[doi] = r
            else:
                ex = by_doi[doi]
                if len(r.get("abstract") or "") > len(ex.get("abstract") or ""):
                    ex["abstract"] = r["abstract"]
                if not ex.get("venue"):
                    ex["venue"] = r.get("venue", "")
                if not ex.get("authors"):
                    ex["authors"] = r.get("authors", [])
                ex["cited_by"] = max(ex.get("cited_by", 0) or 0, r.get("cited_by", 0) or 0)
    return list(by_doi.values())


@st.cache_data(show_spinner=False)
def verify_doi(doi):
    if not doi:
        return False
    headers = {"User-Agent": f"ResearchCoPilot/1.0 (mailto:{MAILTO})"}
    try:
        r = requests.get(f"https://api.crossref.org/works/{doi}", headers=headers, timeout=20)
        return r.status_code == 200
    except requests.RequestException:
        return False


def apa_author(names):
    fams = [n.split()[-1] for n in names if n]
    if not fams:
        return "Anonim"
    if len(fams) == 1:
        return fams[0]
    if len(fams) == 2:
        return f"{fams[0]} & {fams[1]}"
    return f"{fams[0]} et al."


def daftar_pustaka(refs):
    items = []
    for r in sorted(refs, key=lambda x: (x["authors"][0].split()[-1]
                                         if x["authors"] else "z")):
        authors = ", ".join(r["authors"]) if r["authors"] else "Anonim"
        year = r["year"] or "t.t."
        line = f"{authors} ({year}). {r['title']}. {r['venue']}."
        if r["doi"]:
            line += f" https://doi.org/{r['doi']}"
        items.append(line)
    return "\n\n".join(items)


# ============ CLAUDE ============
SYSTEM_PROMPT = """Kamu asisten penulisan akademik untuk mahasiswa Indonesia.
ATURAN WAJIB:
1. Untuk SITASI, gunakan HANYA sumber yang diberikan pengguna. Tulis sitasi
   dalam teks sebagai (NamaBelakang, Tahun) memakai data dari daftar itu.
   JANGAN mengarang sumber, penulis, atau tahun di luar daftar.
2. JANGAN mengarang data lapangan, angka statistik, atau hasil penelitian.
3. Tulis prosa akademik Indonesia yang baku, jelas, dan mengalir.
Ini draf berbantuan AI yang akan dikembangkan mahasiswa; ide & analisis inti
tetap milik mahasiswa."""


def _call_claude(user_prompt, max_tokens=3500):
    import anthropic
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return None
    client = anthropic.Anthropic(api_key=key)
    msg = client.messages.create(
        model="claude-sonnet-4-6",   # ganti bila error 'model not found'
        max_tokens=max_tokens,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return msg.content[0].text


def brainstorm_judul(topik, jenis, metode, n_var, bidang, lokasi):
    lok = f" Lokasi penelitian: {lokasi}." if lokasi.strip() else ""
    return _call_claude(
        f"Bantu rumuskan 10 alternatif JUDUL penelitian yang layak dan dapat "
        f"diteliti untuk:\n- Topik umum: {topik}\n- Jenis: {jenis}\n"
        f"- Metode: {metode}\n- Jumlah variabel bebas (X): {n_var} "
        f"(1 variabel terikat Y disertakan)\n- Bidang ilmu: {bidang}.{lok}\n\n"
        "Syarat: judul spesifik (tidak terlalu luas), sesuai metode, realistis "
        "dikerjakan mahasiswa. Beri nomor 1–10; tiap judul beri satu baris singkat "
        "tentang fokus/variabelnya. Ini alternatif untuk dipilih & disunting sendiri; "
        "judul final wajib disetujui dosen pembimbing.", 1600)


def _refs_block(refs, limit=10):
    blocks = []
    for r in refs[:limit]:
        cite = apa_author(r["authors"])
        abstract = (r["abstract"] or "")[:1500]
        blocks.append(
            f"- Sitasi: ({cite}, {r['year']}) | {r['title']} | {r['venue']}\n"
            f"  Abstrak: {abstract if abstract else '(abstrak tidak tersedia)'}"
        )
    return "\n".join(blocks)


def gen_bab1(refs, topic):
    return _call_claude(
        f"Judul penelitian: {topic}\n\n"
        f"Sumber terverifikasi (gunakan HANYA ini untuk sitasi):\n{_refs_block(refs)}\n\n"
        "Susun draf BAB I PENDAHULUAN yang lengkap dan mengalir:\n"
        "1.1 Latar Belakang (pola corong; sitasi (Penulis, Tahun) dari sumber).\n"
        "1.2 Rumusan Masalah.\n1.3 Tujuan Penelitian.\n"
        "1.4 Manfaat Penelitian (teoretis & praktis).\n"
        "Jangan mengarang angka lapangan.", 3500)


def gen_bab2(refs, topic):
    return _call_claude(
        f"Judul penelitian: {topic}\n\n"
        f"Sumber terverifikasi (gunakan HANYA ini untuk sitasi):\n{_refs_block(refs)}\n\n"
        "Susun draf BAB II KAJIAN PUSTAKA yang mengalir:\n"
        "2.1 Kajian Teori (konsep/variabel utama, dengan sitasi).\n"
        "2.2 Penelitian Terdahulu (sintesis & bandingkan, tunjukkan celah).\n"
        "2.3 Kerangka Berpikir.\n"
        "Gunakan sitasi (Penulis, Tahun) hanya dari sumber di atas.", 3500)


def gen_bab3(refs, topic, metode):
    return _call_claude(
        f"Judul penelitian: {topic}\nMetode: {metode}\n\n"
        f"Sumber (boleh dirujuk bila relevan):\n{_refs_block(refs, limit=6)}\n\n"
        "Susun draf kerangka BAB III METODE PENELITIAN:\n"
        "3.1 Desain/Jenis Penelitian.\n3.2 Populasi dan Sampel.\n"
        "3.3 Instrumen Penelitian.\n3.4 Teknik Pengumpulan Data.\n"
        "3.5 Teknik Analisis Data.\n"
        "Jelaskan pilihan yang lazim untuk desain ini. Untuk keputusan spesifik "
        "(jumlah sampel nyata, lokasi), sebutkan singkat apa yang harus peneliti "
        "tentukan sendiri — tanpa mengarang angka.", 3000)


def to_docx(topic, chapters):
    from docx import Document
    doc = Document()
    doc.add_heading(topic.upper(), level=0)
    for title, body in chapters.items():
        if not body:
            continue
        doc.add_heading(title, level=1)
        for para in (body or "").split("\n"):
            if para.strip():
                doc.add_paragraph(para.strip())
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def show_generated(state_key, extra_note=""):
    val = st.session_state.get(state_key)
    if val == "__NOKEY__":
        st.warning("ANTHROPIC_API_KEY belum diatur di Railway.")
    elif isinstance(val, str) and val.startswith("__ERROR__"):
        st.error(f"Error: {val[9:]}")
    elif val:
        st.markdown(val)
        if extra_note:
            st.caption(extra_note)


def run_module(func, *args, state_key=""):
    with st.spinner("Menyusun draf..."):
        try:
            res = func(*args)
            st.session_state[state_key] = res if res is not None else "__NOKEY__"
        except Exception as e:
            st.session_state[state_key] = f"__ERROR__{e}"


# ============ TAMPILAN ============
st.title("📚 Research Co-Pilot")
st.caption("Draf dari referensi asli terverifikasi. Sitasi nyata, bukan karangan.")

# --- Brainstorming Judul (panel atas) ---
with st.expander("💡 Belum punya judul? Brainstorming Judul"):
    bt = st.text_area("Topik umum", key="bt_topik",
                      placeholder="mis. model pembelajaran kooperatif pada PJOK")
    c1, c2, c3 = st.columns(3)
    with c1:
        b_jenis = st.selectbox("Jenis", ["Skripsi", "Tesis", "Disertasi"])
    with c2:
        b_metode = st.selectbox("Metode", ["Kuantitatif", "Kualitatif",
                                           "Studi Pustaka", "PTK", "R&D",
                                           "Mixed Method", "Eksperimen"])
    with c3:
        b_var = st.selectbox("Jumlah variabel X", ["1", "2", "3", "4"], index=1)
    c4, c5 = st.columns(2)
    with c4:
        b_bidang = st.text_input("Bidang ilmu", value="Pendidikan Olahraga")
    with c5:
        b_lokasi = st.text_input("Lokasi (opsional)", placeholder="mis. SD di Bandung")
    if st.button("✨ Generate Judul") and bt.strip():
        run_module(brainstorm_judul, bt, b_jenis, b_metode, b_var, b_bidang,
                   b_lokasi, state_key="judul_options")
    show_generated("judul_options",
                   "ℹ️ Pilih & sunting satu judul, salin ke kolom di bawah. "
                   "Judul final wajib disetujui dosen pembimbing.")

st.divider()

# --- Pencarian referensi ---
topic = st.text_input("Judul / topik penelitian",
                      placeholder="mis. Perbandingan model Jigsaw dan TGT ...")
col1, col2 = st.columns(2)
with col1:
    n = st.slider("Jumlah referensi", 5, 20, 10)
with col2:
    metode = st.text_input("Metode (untuk Bab III)", value="Kuantitatif kuasi-eksperimen")

if st.button("🔎 Cari Referensi Terverifikasi", type="primary") and topic.strip():
    with st.spinner("Mencari & memverifikasi DOI..."):
        cand = merge_dedupe(search_openalex(topic, n * 2),
                            search_semantic_scholar(topic, n * 2))
        verified = []
        for c in cand:
            if verify_doi(c["doi"]):
                verified.append(c)
            if len(verified) >= n:
                break
    st.session_state["refs"] = verified
    st.session_state["topic"] = topic
    for k in ("bab1", "bab2", "bab3"):
        st.session_state.pop(k, None)

refs = st.session_state.get("refs", [])
if refs:
    st.success(f"{len(refs)} referensi terverifikasi ✅ (DOI dikonfirmasi di Crossref)")
    with st.expander("Lihat daftar referensi"):
        for r in refs:
            st.markdown(f"- **{r['title']}** — {apa_author(r['authors'])} "
                        f"({r['year']}) · [🔗]({'https://doi.org/'+r['doi']})")

    st.subheader("📝 Susun Draf Bab")
    tabs = st.tabs(["BAB I — Pendahuluan", "BAB II — Kajian Pustaka", "BAB III — Metode"])
    jobs = [("bab1", gen_bab1, (refs, topic)),
            ("bab2", gen_bab2, (refs, topic)),
            ("bab3", gen_bab3, (refs, topic, metode))]
    note = ("ℹ️ Draf berbantuan AI dari referensi terverifikasi. Kembangkan dengan "
            "analisis Anda, periksa tiap sitasi sesuai isinya, nyatakan penggunaan AI.")
    for tab, (key, fn, args) in zip(tabs, jobs):
        with tab:
            if st.button(f"Buat draf {key.upper()}", key=f"btn_{key}"):
                run_module(fn, *args, state_key=key)
            show_generated(key, note)

    chapters = {
        "BAB I PENDAHULUAN": st.session_state.get("bab1"),
        "BAB II KAJIAN PUSTAKA": st.session_state.get("bab2"),
        "BAB III METODE PENELITIAN": st.session_state.get("bab3"),
    }
    chapters = {k: v for k, v in chapters.items() if v and not str(v).startswith("__")}
    if chapters:
        chapters["DAFTAR PUSTAKA"] = daftar_pustaka(refs)
        st.download_button(
            "⬇️ Unduh Word (bab yang sudah dibuat + daftar pustaka)",
            data=to_docx(topic, chapters),
            file_name="draf_skripsi.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.info("BAB IV (Hasil) disusun dari DATA penelitian Anda sendiri, dan BAB V "
            "dari temuan Anda — tidak dibuat otomatis, karena mengarang hasil "
            "melanggar integritas akademik.")
