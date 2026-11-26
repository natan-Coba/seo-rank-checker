import io
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="SEO Rank Checker", layout="centered")

st.title("SEO Rank Checker")
st.write("Cek posisi ranking domain target secara organik maupun iklan via SearchAPI.io dan ekspor ke Excel.")

daftar_bulan = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

topik = st.text_input("Topik Utama", "Mobil Toyota")
bulan = st.selectbox("Bulan Pengecekan", daftar_bulan, index=10)
target_domain = st.text_input("Domain Target", "auto2000.co.id")
lokasi_pencarian = st.text_input("Lokasi Pencarian (opsional, contoh: Jakarta,Indonesia)", "Jakarta,Indonesia")
api_key = st.text_input("SearchAPI.io API Key", type="password")

raw_keywords = st.text_area(
    "Daftar Keyword (satu baris per keyword)",
    placeholder="land cruiser fj60\ngenerasi alphard\ngenerasi camry"
)

debug_mode = st.checkbox("Tampilkan Debug JSON (keyword terakhir)")


def bersihkan_domain(domain_input: str) -> str:
    cleaned = domain_input.strip().lower()
    for prefix in ("https://", "http://", "www."):
        if cleaned.startswith(prefix):
            cleaned = cleaned[len(prefix):]
    return cleaned.split("/")[0]


def cek_ranking_searchapi(keyword: str, domain_bersih: str, key: str, location: str = ""):
    url = "https://www.searchapi.io/api/v1/search"
    params = {
        "engine": "google",
        "q": keyword,
        "gl": "id",
        "hl": "id",
        "api_key": key,
    }
    if location.strip():
        params["location"] = location.strip()

    try:
        resp = requests.get(url, params=params, timeout=30)
        data = resp.json()

        if "error" in data:
            return "API Error", "-", f"Error: {data['error']}", data

        # 1. Cek hasil organik
        for idx, item in enumerate(data.get("organic_results", []), start=1):
            link = item.get("link", "")
            item_domain = item.get("domain", "").lower()
            if domain_bersih in link.lower() or domain_bersih in item_domain:
                pos = item.get("position", idx)
                return "Organik", pos, link or f"https://{domain_bersih}", data

        # 2. Cek iklan (ads) jika tidak muncul di organik
        for item in data.get("ads", []):
            link = item.get("link", "")
            displayed = item.get("displayed_link", "")
            if domain_bersih in link.lower() or domain_bersih in displayed.lower():
                return "Sponsored", "Ads", link or displayed or "-", data

        return "Tidak Ditemukan", "> 100", "-", data

    except Exception as exc:
        return "Error", "-", str(exc), {}


if st.button("Mulai Pengecekan"):
    keywords = [line.strip() for line in raw_keywords.splitlines() if line.strip()]
    domain_bersih = bersihkan_domain(target_domain)

    if not api_key.strip():
        st.error("Masukkan API Key SearchAPI.io terlebih dahulu.")
    elif not domain_bersih:
        st.error("Masukkan domain target yang valid.")
    elif not keywords:
        st.warning("Masukkan minimal satu keyword.")
    else:
        st.info(f"Memproses {len(keywords)} keyword untuk bulan {bulan}...")
        progress = st.progress(0)
        rows = []
        last_json = {}

        for i, kw in enumerate(keywords, start=1):
            tipe, rank, url_found, raw_json = cek_ranking_searchapi(
                kw, domain_bersih, api_key.strip(), lokasi_pencarian
            )
            last_json = raw_json
            rows.append({
                "Bulan": bulan,
                "Topik": topik,
                "Keyword": kw,
                "Domain Target": domain_bersih,
                "Lokasi": lokasi_pencarian.strip() or "Default (gl=id)",
                "Tipe Tampil": tipe,
                "Rank": rank,
                "URL / Keterangan": url_found,
            })
            progress.progress(i / len(keywords))

        df = pd.DataFrame(rows)
        st.success("Pengecekan selesai.")
        st.dataframe(df, use_container_width=True)

        if debug_mode:
            st.subheader("Debug JSON (Keyword Terakhir)")
            st.json(last_json)

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=f"Ranking {bulan}"[:31])
        excel_buffer.seek(0)

        safe_topik = "_".join(topik.lower().split()) or "rekap"
        filename = f"ranking_{safe_topik}_{bulan.lower()}.xlsx"

        st.download_button(
            label=f"Download Rekap {bulan} (.xlsx)",
            data=excel_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )