"""
Research Co-Pilot — Mesin Referensi Terverifikasi (MVP)
--------------------------------------------------------
Menampilkan HANYA referensi yang benar-benar ada:
  1. Cari kandidat di OpenAlex
  2. Verifikasi tiap DOI ke Crossref (yang gagal disaring)
  3. Tampilkan + ekspor RIS untuk Zotero/Mendeley

TIDAK menulis paper. TIDAK mengarang data/statistik/referensi.
Ide, data, analisis, dan tulisan tetap milik & tanggung jawab peneliti.

Deploy: taruh file ini + requirements.txt di repo GitHub, hubungkan ke Railway.
requirements.txt:
    streamlit
    requests
"""

import streamlit as st
import requests

# --- Konfigurasi ---
MAILTO = "bayushadega74.com"   # GANTI dengan email Anda (untuk "polite pool")
OPENALEX_KEY = None                # OpenAlex kini bisa minta API key gratis;
                                   # isi di sini jika dapat error otorisasi.

st.set_page_config(page_title="Research Co-Pilot — Referensi Terverifikasi",
                   page_icon="📚")


# --- Fungsi inti ---
@st.cache_data(show_spinner=False)
def search_openalex(query, per_page=20):
    """Cari paper NYATA dari OpenAlex."""
    params = {"search": query, "per-page": per_page, "mailto": MAILTO}
    if OPENALEX_KEY:
        params["api_key"] = OPENALEX_KEY
    r = requests.get("https://api.openalex.org/works", params=params, timeout=25)
    r.raise_for_status()
    out = []
    for w in r.json().get("results", []):
        doi = (w.get("doi") or "").replace("https://doi.org/", "")
        src = (w.get("primary_location") or {}).get("source") or {}
        out.append({
            "title": w.get("title") or "(tanpa judul)",
            "year": w.get("publication_year"),
            "doi": doi,
            "authors": [a["author"]["display_name"]
                        for a in w.get("authorships", [])][:5],
            "venue": src.get("display_name") or "",
            "cited_by": w.get("cited_by_count", 0),
        })
    return out


@st.cache_data(show_spinner=False)
def verify_doi(doi):
    """True jika DOI benar-benar terdaftar di Crossref."""
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
    """Ubah daftar referensi jadi format RIS (untuk Zotero/Mendeley)."""
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


# --- Tampilan ---
st.title("📚 Research Co-Pilot")
st.caption("Referensi asli terverifikasi — bukan karangan AI. "
           "Kamu tetap penulis karyamu.")

query = st.text_input("Topik penelitian Anda",
                      placeholder="mis. HIIT VO2max adolescent athletes")
n = st.slider("Jumlah referensi", 5, 20, 10)

if st.button("🔎 Cari Referensi Terverifikasi", type="primary") and query.strip():
    with st.spinner("Mencari di OpenAlex & memverifikasi DOI di Crossref..."):
        candidates = search_openalex(query, per_page=n * 2)
        verified = []
        for c in candidates:
            if verify_doi(c["doi"]):
                verified.append(c)
            if len(verified) >= n:
                break

    if not verified:
        st.warning("Tidak ada referensi ber-DOI terverifikasi untuk topik ini. "
                   "Coba kata kunci lain.")
    else:
        st.success(f"{len(verified)} referensi terverifikasi ✅ "
                   f"(DOI dikonfirmasi ada di Crossref)")
        for r in verified:
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
            data=to_ris(verified),
            file_name="referensi.ris",
            mime="application/x-research-info-systems",
        )

st.info("Tool ini hanya menampilkan referensi yang benar-benar ada. "
        "Ide, data, analisis, dan tulisan tetap milik dan tanggung jawab Anda.")
