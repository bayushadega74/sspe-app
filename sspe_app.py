"""
Research Co-Pilot (v2 — dua sumber pencarian)
---------------------------------------------
1. Cari referensi NYATA dari OpenAlex + Semantic Scholar
   -> gabung, buang duplikat (berdasarkan DOI), pilih abstrak terpanjang
   -> verifikasi tiap DOI ke Crossref (yang gagal disaring)
2. Bantu sintesis 'Penelitian Terdahulu' — grounded HANYA pada abstrak terverifikasi.

Ide, data, analisis, dan tulisan tetap milik & tanggung jawab peneliti.

PENTING (keamanan): API key TIDAK ditaruh di kode ini (repo Anda publik).
Simpan sebagai environment variable di Railway:
  - ANTHROPIC_API_KEY (wajib, untuk modul sintesis)
  - SEMANTIC_SCHOLAR_API_KEY (OPSIONAL, hanya kalau butuh rate lebih tinggi)

requirements.txt:
    streamlit
    requests
    anthropic
"""

import os
import requests
import streamlit as st

# --- Konfigurasi ---
MAILTO = "bayushadega74@gmail.com"   # email Anda (untuk polite pool)
OPENALEX_KEY = None                  # isi bila OpenAlex minta API key gratis

st.set_page_config(page_title="Research Co-Pilot", page_icon="📚")


# ============ BAGIAN 1: SUMBER PENCARIAN ============
def reconstruct_abstract(inv):
    """OpenAlex menyimpan abstrak sebagai indeks terbalik; kita susun ulang."""
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
        r = requests.get("https://api.openalex.org/works",
                         params=params, timeout=25)
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
    """Sumber kedua. Gagal dengan aman (kembali []) bila rate-limit/error."""
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
    """Gabung beberapa sumber, buang duplikat berdasarkan DOI,
    pilih abstrak yang lebih panjang saat sama."""
    by_doi = {}
    for lst in lists:
        for r in lst:
            doi = (r.get("doi") or "").strip().lower()
            if not doi:
                continue  # tanpa DOI tak bisa diverifikasi; lewati
            if doi not in by_doi:
                by_doi[doi] = r
            else:
                existing = by_doi[doi]
                # pilih abstrak yang lebih panjang
                if len(r.get("abstract") or "") > len(existing.get("abstract") or ""):
                    existing["abstract"] = r["abstract"]
                # lengkapi field yang kosong
                if not existing.get("venue"):
                    existing["venue"] = r.get("venue", "")
                if not existing.get("authors"):
                    existing["authors"] = r.get("authors", [])
                existing["cited_by"] = max(existing.get("cited_by", 0) or 0,
                                           r.get("cited_by", 0) or 0)
    return list(by_doi.values())


@st.cache_data(show_spinner=False)
def verify_doi(doi):
    if not doi:
        return False
    headers = {"User-Agent": f"ResearchCoPilot/1.0 (mailto:{MAILTO})"}
    try:
        r = requests.get(f"https://api.crossref.org/works/{doi}",
                         headers=headers, timeout=20)
        return r.status_code == 200
    except requests.RequestException:
        return False


def to_ris(refs):
    lines = []
    for r in refs:
        lines.append("TY  - JOUR")
        for a in r["authors"]:
            lines.append(f"AU  - {a}")
        lines.append(f"TI  - {r['title']}")
        if r["year"]:
            lines.append(f"PY  - {r['year']}")
        if r["venue"]:
            lines.append(f"JO  - {r['venue']}")
        if r["doi"]:
            lines.append(f"DO  - {r['doi']}")
        lines.append("ER  - ")
        lines.append("")
    return "\n".join(lines)


# ============ BAGIAN 2: MODUL PENULISAN (Claude, grounded) ============
SYSTEM_PROMPT = """Kamu asisten riset akademik untuk mahasiswa Indonesia.
ATURAN WAJIB:
1. JANGAN mengarang referensi, data, atau temuan. Gunakan HANYA sumber yang
   diberikan pengguna. Jika sebuah abstrak tidak menyebut temuan tertentu,
   tulis "tidak disebutkan di abstrak" — jangan menebak.
2. Hasilmu adalah DRAF KERJA/kerangka, BUKAN naskah final. Sisipkan penanda
   [KEMBANGKAN DENGAN KALIMAT ANDA] pada bagian yang harus ditulis mahasiswa.
3. Akhiri dengan 2-3 pertanyaan reflektif.
4. Jangan pernah membantu menghindari deteksi plagiarisme.
5. Bahasa Indonesia akademik baku, sitasi APA. Ide inti & analisis milik mahasiswa."""


def refs_to_context(refs, limit=8):
    blocks = []
    for r in refs[:limit]:
        authors = ", ".join(r["authors"])
        abstract = (r["abstract"] or "")[:2000]
        blocks.append(
            f"- {authors} ({r['year']}). {r['title']}. {r['venue']}. "
            f"DOI: {r['doi']}\n"
            f"  Abstrak: {abstract if abstract else '(abstrak tidak tersedia)'}"
        )
    return "\n".join(blocks)


def synthesize_prior_research(refs, topic):
    import anthropic
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return None
    client = anthropic.Anthropic(api_key=key)
    user_prompt = (
        f"Topik penelitian: {topic}\n\n"
        "Gunakan HANYA sumber terverifikasi berikut "
        "(jangan menambah dari ingatanmu):\n"
        f"{refs_to_context(refs)}\n\n"
        "Buat DRAF KERJA sintesis 'Penelitian Terdahulu':\n"
        "1. Tabel: Penulis (tahun) | Fokus/Metode | Temuan utama | Celah — "
        "isi HANYA dari abstrak di atas.\n"
        "2. Satu paragraf sintesis yang menunjukkan pola dan celah, dengan "
        "sitasi APA, sisipkan [KEMBANGKAN DENGAN KALIMAT ANDA].\n"
        "Jangan mengarang temuan yang tidak ada di abstrak."
    )
    msg = client.messages.create(
        model="claude-sonnet-4-6",   # jika error 'model not found',
                                     # ganti ke model terbaru dari console.anthropic.com
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return msg.content[0].text


# ============ TAMPILAN ============
st.title("📚 Research Co-Pilot")
st.caption("Referensi asli terverifikasi — bukan karangan AI. "
           "Kamu tetap penulis karyamu.")

topic = st.text_input("Topik penelitian Anda",
                      placeholder="mis. HIIT VO2max adolescent athletes")
n = st.slider("Jumlah referensi", 5, 20, 10)

if st.button("🔎 Cari Referensi Terverifikasi", type="primary") and topic.strip():
    with st.spinner("Mencari di OpenAlex + Semantic Scholar, "
                    "lalu memverifikasi DOI di Crossref..."):
        cand_oa = search_openalex(topic, per_page=n * 2)
        cand_ss = search_semantic_scholar(topic, limit=n * 2)
        candidates = merge_dedupe(cand_oa, cand_ss)
        verified = []
        for c in candidates:
            if verify_doi(c["doi"]):
                verified.append(c)
            if len(verified) >= n:
                break
    st.session_state["refs"] = verified
    st.session_state["topic"] = topic
    st.session_state["n_sources"] = f"OpenAlex: {len(cand_oa)} · Semantic Scholar: {len(cand_ss)}"

refs = st.session_state.get("refs", [])
if refs:
    st.success(f"{len(refs)} referensi terverifikasi ✅ "
               f"(DOI dikonfirmasi ada di Crossref)")
    if st.session_state.get("n_sources"):
        st.caption(f"Kandidat dari — {st.session_state['n_sources']}")
    for r in refs:
        authors = ", ".join(r["authors"])
        if len(r["authors"]) >= 5:
            authors += " dkk."
        st.markdown(
            f"**{r['title']}**  \n"
            f"{authors} ({r['year']}) — *{r['venue']}*  \n"
            f"Sitasi: {r['cited_by']} · "
            f"[🔗 Buka artikel](https://doi.org/{r['doi']})"
        )
        st.divider()

    st.download_button(
        "⬇️ Unduh semua (RIS untuk Zotero/Mendeley)",
        data=to_ris(refs),
        file_name="referensi.ris",
        mime="application/x-research-info-systems",
    )

    st.subheader("✍️ Bantu Sintesis Penelitian Terdahulu (draf kerja)")
    st.caption("AI menyusun KERANGKA dari abstrak referensi di atas. "
               "Anda tetap menulis analisis akhirnya sendiri.")
    if st.button("Buat draf sintesis"):
        with st.spinner("Menyusun draf dari referensi terverifikasi..."):
            try:
                result = synthesize_prior_research(
                    refs, st.session_state.get("topic", topic))
            except Exception as e:
                st.error(f"Terjadi error saat memanggil AI: {e}")
                result = "ERROR"
        if result is None:
            st.warning("API key Anthropic belum diatur. Tambahkan variabel "
                       "ANTHROPIC_API_KEY di Railway.")
        elif result != "ERROR":
            st.markdown(result)
            st.caption("⚠️ Ini draf kerja. Tulis ulang dengan kalimat Anda, "
                       "verifikasi tiap temuan, dan nyatakan penggunaan AI.")

st.info("Tool ini menampilkan referensi yang benar-benar ada dan membantu "
        "menyusun KERANGKA. Ide, data, analisis, dan tulisan tetap milik "
        "dan tanggung jawab Anda.")
