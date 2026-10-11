
import re
import csv
from io import BytesIO, StringIO
from datetime import datetime

import streamlit as st
from docx import Document
from pypdf import PdfReader


# ============================================================
# ASWAJA REF-AI V0.2
# Asisten Akademik Berbasis Generative AI dengan Sistem
# Validasi Referensi Aswaja An-Nahdliyah dan Sitasi Otomatis
# Karya Ilmiah Mahasiswa
#
# Fokus tunggal:
# Literatur dan Referensi Aswaja An-Nahdliyah
#
# V0.2:
# - Lima menu utama dengan submenu
# - Unggah PDF dan DOCX
# - Identitas dokumen
# - Deteksi awal kutipan penulis-tahun
# - Deteksi awal daftar pustaka
# - Laporan hasil deteksi awal
# - Dokumentasi pengembangan R&D
#
# Belum melakukan validasi sumber asli berbasis AI.
# ============================================================

st.set_page_config(
    page_title="ASWAJA REF-AI",
    page_icon="📚",
    layout="wide"
)

JUDUL = (
    "Asisten Akademik Berbasis Generative AI dengan Sistem "
    "Validasi Referensi Aswaja An-Nahdliyah dan Sitasi "
    "Otomatis Karya Ilmiah Mahasiswa"
)

VERSI = "0.2"


# ============================================================
# 1. SESSION STATE
# ============================================================

DEFAULTS = {
    "isi_dokumen": "",
    "nama_dokumen": "",
    "identitas": {},
    "kutipan": [],
    "daftar_pustaka": [],
    "hasil_baca": False,
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# 2. FUNGSI PEMBACA DOKUMEN
# ============================================================

def baca_dokumen(file):
    data = BytesIO(file.getvalue())
    nama = file.name.lower()

    if nama.endswith(".docx"):
        dokumen = Document(data)

        bagian = []

        for paragraf in dokumen.paragraphs:
            if paragraf.text.strip():
                bagian.append(paragraf.text)

        for tabel in dokumen.tables:
            for baris in tabel.rows:
                isi_baris = [
                    cell.text.strip()
                    for cell in baris.cells
                ]
                bagian.append(" | ".join(isi_baris))

        return "\n".join(bagian)

    if nama.endswith(".pdf"):
        pdf = PdfReader(data)

        bagian = []
        for halaman in pdf.pages:
            teks = halaman.extract_text() or ""
            bagian.append(teks)

        return "\n".join(bagian)

    raise ValueError("Format file tidak didukung.")


# ============================================================
# 3. DETEKSI AWAL KUTIPAN
# ============================================================

def temukan_kutipan(teks):
    pola = (
        r"\([^()\n]{2,90}?,\s*"
        r"(?:19|20)\d{2}"
        r"(?:\s*[:,]\s*[^()\n]{1,25})?\)"
    )

    hasil = re.findall(pola, teks)

    # Menampilkan kutipan unik.
    # Jumlah kemunculan sebenarnya belum dihitung.
    return list(dict.fromkeys(hasil))


# ============================================================
# 4. DETEKSI AWAL DAFTAR PUSTAKA
# ============================================================

def temukan_daftar_pustaka(teks):
    pola_judul = re.compile(
        r"(?im)^\s*(?:"
        r"daftar pustaka|bibliografi|references"
        r")\s*$"
    )

    cocok = list(pola_judul.finditer(teks))

    if not cocok:
        return []

    bagian = teks[cocok[-1].end():]

    # Deteksi awal, belum merupakan parser bibliografi.
    pola_akhir = re.compile(
        r"(?im)^\s*(?:lampiran|appendix)\s*$"
    )

    akhir = pola_akhir.search(bagian)

    if akhir:
        bagian = bagian[:akhir.start()]

    baris = [
        b.strip()
        for b in bagian.splitlines()
        if b.strip()
    ]

    return baris[:500]


# ============================================================
# 5. LAPORAN DETEKSI AWAL
# ============================================================

def buat_laporan():
    nama = st.session_state.nama_dokumen
    teks = st.session_state.isi_dokumen
    kutipan = st.session_state.kutipan
    pustaka = st.session_state.daftar_pustaka

    isi = [
        "LAPORAN DETEKSI AWAL",
        "ASWAJA REF-AI",
        JUDUL,
        "",
        f"Versi aplikasi: {VERSI}",
        f"Tanggal: {datetime.now():%d-%m-%Y}",
        f"Dokumen: {nama or 'Belum tersedia'}",
        f"Jumlah kata: {len(teks.split())}",
        f"Jumlah kutipan unik terdeteksi: {len(kutipan)}",
        f"Jumlah baris daftar pustaka: {len(pustaka)}",
        "",
        "KUTIPAN TERDETEKSI",
    ]

    for nomor, item in enumerate(kutipan, 1):
        isi.append(f"{nomor}. {item}")

    isi.extend([
        "",
        "BARIS DAFTAR PUSTAKA TERDETEKSI",
    ])

    for nomor, item in enumerate(pustaka, 1):
        isi.append(f"{nomor}. {item}")

    isi.extend([
        "",
        "CATATAN PENTING",
        "Laporan ini hanya hasil deteksi awal.",
        "Belum ada pemeriksaan keaslian sumber.",
        "Belum ada validasi kesesuaian isi kutipan.",
        "Belum ada verifikasi nomor halaman.",
        "Belum ada klasifikasi otomatis sumber Aswaja An-Nahdliyah.",
        "Belum ada validasi berbasis Generative AI.",
        "Hasil ini tidak boleh dianggap sebagai bukti",
        "bahwa suatu referensi valid atau tidak valid.",
    ])

    return "\n".join(isi)


# ============================================================
# 6. MENU UTAMA DAN SUBMENU
# ============================================================

MENU = {
    "🏠 Beranda": [
        "Dashboard",
        "Panduan Penggunaan",
        "Informasi Aplikasi",
    ],

    "📤 Unggah Karya Ilmiah": [
        "Unggah Dokumen",
        "Identitas Dokumen",
        "Hasil Pembacaan",
        "Deteksi Kutipan",
        "Deteksi Daftar Pustaka",
    ],

    "🔎 Validasi Literatur dan Referensi Aswaja An-Nahdliyah": [
        "Validasi Keaslian Literatur",
        "Validasi Kesesuaian Kutipan",
        "Pemeriksaan Halaman Kutipan",
        "Validasi Bibliografi",
        "Pencarian Referensi Pengganti",
        "Hasil Validasi",
    ],

    "📝 Sitasi Otomatis": [
        "Pilihan Gaya Sitasi",
        "Perbaikan Kutipan",
        "Perbaikan Catatan Kaki",
        "Perbaikan Daftar Pustaka",
        "Pratinjau Hasil",
    ],

    "📊 Laporan Validasi": [
        "Ringkasan Hasil",
        "Temuan Kesalahan",
        "Rekomendasi Perbaikan",
        "Unduh Laporan",
        "Dokumentasi Pengembangan R&D",
    ],
}


# ============================================================
# 7. HEADER
# ============================================================

st.title("📚 ASWAJA REF-AI")

st.caption(JUDUL)

st.info(
    "Versi 0.2: Pengembangan menu, submenu, pembacaan "
    "dokumen, dan deteksi awal. Mesin validasi sumber asli "
    "berbasis Generative AI serta sitasi otomatis masih "
    "dalam tahap pengembangan."
)


# ============================================================
# 8. SIDEBAR
# ============================================================

with st.sidebar:
    st.header("📚 ASWAJA REF-AI")

    menu = st.radio(
        "Menu Utama",
        list(MENU.keys()),
        key="menu_utama"
    )

    st.divider()

    st.subheader("Submenu")

    submenu = st.radio(
        "Pilih Submenu",
        MENU[menu],
        key=f"submenu_{menu}"
    )

    st.divider()

    st.caption(
        "Penelitian R&D Model ADDIE | "
        "Literatur dan Referensi Aswaja An-Nahdliyah"
    )


# ============================================================
# 9. BERANDA
# ============================================================

if menu == "🏠 Beranda":

    if submenu == "Dashboard":

        st.subheader("Dashboard ASWAJA REF-AI")

        a, b, c = st.columns(3)

        a.metric(
            "Dokumen Aktif",
            1 if st.session_state.hasil_baca else 0
        )

        b.metric(
            "Kutipan Unik Terdeteksi",
            len(st.session_state.kutipan)
        )

        c.metric(
            "Baris Daftar Pustaka",
            len(st.session_state.daftar_pustaka)
        )

        st.write(
            "ASWAJA REF-AI dikembangkan untuk membantu "
            "mahasiswa melakukan validasi literatur dan "
            "referensi Aswaja An-Nahdliyah serta "
            "penyusunan sitasi otomatis."
        )

        st.warning(
            "Angka di dashboard merupakan hasil deteksi awal, "
            "bukan hasil validasi referensi."
        )

    elif submenu == "Panduan Penggunaan":

        st.subheader("Panduan Penggunaan")

        st.markdown("""
        1. Buka menu **Unggah Karya Ilmiah**.
        2. Unggah dokumen PDF atau DOCX.
        3. Periksa hasil pembacaan dokumen.
        4. Lihat kutipan dan daftar pustaka terdeteksi.
        5. Gunakan laporan deteksi awal.

        **Tahap pengembangan berikutnya:**
        validasi otomatis referensi Aswaja An-Nahdliyah,
        pemeriksaan kutipan terhadap sumber asli,
        rekomendasi referensi, dan sitasi otomatis.
        """)

    elif submenu == "Informasi Aplikasi":

        st.subheader("Informasi Aplikasi")

        st.write("**Nama:** ASWAJA REF-AI")
        st.write("**Versi:**", VERSI)
        st.write("**Teknologi:** Python dan Streamlit")
        st.write("**Model penelitian:** R&D ADDIE")
        st.write("**Judul penelitian:**", JUDUL)

        st.write(
            "**Fokus:** Validasi Literatur dan Referensi "
            "Aswaja An-Nahdliyah."
        )


# ============================================================
# 10. UNGGAH KARYA ILMIAH
# ============================================================

elif menu == "📤 Unggah Karya Ilmiah":

    if submenu == "Unggah Dokumen":

        st.subheader("Unggah Karya Ilmiah")

        file = st.file_uploader(
            "Pilih satu file PDF atau DOCX",
            type=["pdf", "docx"]
        )

        if file is not None:

            if st.button(
                "📖 Baca dan Analisis Awal Dokumen",
                type="primary"
            ):

                try:
                    with st.spinner("Membaca dokumen..."):

                        teks = baca_dokumen(file)

                        if not teks.strip():
                            st.error(
                                "Teks dokumen tidak terbaca. "
                                "PDF hasil pemindaian mungkin "
                                "memerlukan OCR."
                            )

                        else:
                            st.session_state.isi_dokumen = teks
                            st.session_state.nama_dokumen = file.name
                            st.session_state.kutipan = (
                                temukan_kutipan(teks)
                            )
                            st.session_state.daftar_pustaka = (
                                temukan_daftar_pustaka(teks)
                            )
                            st.session_state.hasil_baca = True

                            st.success(
                                "Dokumen berhasil dibaca."
                            )

                except Exception as error:
                    st.error(
                        f"Dokumen belum dapat dibaca: {error}"
                    )

        if st.session_state.hasil_baca:
            st.write(
                "**Dokumen aktif:**",
                st.session_state.nama_dokumen
            )

    elif submenu == "Identitas Dokumen":

        st.subheader("Identitas Karya Ilmiah")

        identitas = st.session_state.identitas

        with st.form("form_identitas"):

            nama = st.text_input(
                "Nama Mahasiswa",
                value=identitas.get("nama", "")
            )

            nim = st.text_input(
                "NIM",
                value=identitas.get("nim", "")
            )

            perguruan_tinggi = st.text_input(
                "Perguruan Tinggi",
                value=identitas.get("kampus", "")
            )

            jenjang = st.selectbox(
                "Jenjang Pendidikan",
                ["S1", "S2", "S3"],
                index=["S1", "S2", "S3"].index(
                    identitas.get("jenjang", "S2")
                )
            )

            judul = st.text_area(
                "Judul Karya Ilmiah",
                value=identitas.get("judul", "")
            )

            simpan = st.form_submit_button(
                "💾 Simpan Identitas"
            )

        if simpan:
            st.session_state.identitas = {
                "nama": nama,
                "nim": nim,
                "kampus": perguruan_tinggi,
                "jenjang": jenjang,
                "judul": judul,
            }

            st.success("Identitas disimpan pada sesi ini.")

    elif submenu == "Hasil Pembacaan":

        st.subheader("Hasil Pembacaan Dokumen")

        teks = st.session_state.isi_dokumen

        if teks:

            st.write(
                "**Nama file:**",
                st.session_state.nama_dokumen
            )

            st.metric(
                "Jumlah Kata",
                len(teks.split())
            )

            with st.expander("Lihat Teks Dokumen"):
                st.text_area(
                    "Hasil Ekstraksi",
                    teks,
                    height=350,
                    disabled=True
                )

        else:
            st.info(
                "Belum ada dokumen. "
                "Gunakan submenu Unggah Dokumen."
            )

    elif submenu == "Deteksi Kutipan":

        st.subheader("Deteksi Awal Kutipan")

        kutipan = st.session_state.kutipan

        if not st.session_state.hasil_baca:
            st.info("Unggah dan baca dokumen terlebih dahulu.")

        else:
            st.metric(
                "Jumlah Kutipan Unik",
                len(kutipan)
            )

            if kutipan:
                for nomor, item in enumerate(kutipan, 1):
                    st.write(f"{nomor}. {item}")

            else:
                st.warning(
                    "Belum ditemukan kutipan dengan pola "
                    "penulis-tahun yang didukung."
                )

            st.caption(
                "Deteksi masih terbatas pada pola tertentu. "
                "Catatan kaki dan kutipan naratif belum "
                "teridentifikasi secara lengkap."
            )

    elif submenu == "Deteksi Daftar Pustaka":

        st.subheader("Deteksi Awal Daftar Pustaka")

        pustaka = st.session_state.daftar_pustaka

        if not st.session_state.hasil_baca:
            st.info("Unggah dan baca dokumen terlebih dahulu.")

        elif pustaka:
            st.write(
                "**Baris terdeteksi:**",
                len(pustaka)
            )

            for nomor, item in enumerate(pustaka, 1):
                st.write(f"{nomor}. {item}")

            st.caption(
                "Setiap baris belum tentu merupakan satu "
                "entri bibliografi. Pemeriksaan lanjutan "
                "akan dikembangkan."
            )

        else:
            st.warning(
                "Bagian Daftar Pustaka belum terdeteksi."
            )


# ============================================================
# 11. VALIDASI LITERATUR DAN REFERENSI
#     ASWAJA AN-NAHDLIYAH
# ============================================================

elif menu.startswith("🔎 Validasi Literatur"):

    st.subheader(submenu)

    if st.session_state.hasil_baca:
        st.write(
            "**Dokumen aktif:**",
            st.session_state.nama_dokumen
        )
    else:
        st.info(
            "Unggah karya ilmiah terlebih dahulu."
        )

    penjelasan = {
        "Validasi Keaslian Literatur":
            "Memeriksa keberadaan kitab, buku, jurnal, "
            "dokumen ke-NU-an, dan penelitian terkait "
            "Aswaja An-Nahdliyah.",

        "Validasi Kesesuaian Kutipan":
            "Membandingkan isi kutipan mahasiswa "
            "dengan teks sumber asli yang dapat diakses.",

        "Pemeriksaan Halaman Kutipan":
            "Memeriksa kecocokan nomor halaman "
            "berdasarkan edisi sumber yang tersedia.",

        "Validasi Bibliografi":
            "Memeriksa nama penulis, judul, tahun, "
            "penerbit, edisi, dan metadata sumber.",

        "Pencarian Referensi Pengganti":
            "Menelusuri literatur Aswaja An-Nahdliyah "
            "yang relevan dan dapat diverifikasi.",

        "Hasil Validasi":
            "Menampilkan hasil pemeriksaan sumber, "
            "bukti, serta rekomendasi perbaikan.",
    }

    st.write(penjelasan.get(submenu, ""))

    st.warning(
        "Fitur validasi otomatis pada submenu ini "
        "belum diaktifkan. Belum tersedia bukti "
        "pemeriksaan terhadap sumber asli."
    )

    st.info(
        "Pengembangan berikutnya akan mengintegrasikan "
        "Generative AI, bank literatur Aswaja An-Nahdliyah, "
        "dan mekanisme verifikasi berbasis sumber."
    )


# ============================================================
# 12. SITASI OTOMATIS
# ============================================================

elif menu == "📝 Sitasi Otomatis":

    st.subheader(submenu)

    if submenu == "Pilihan Gaya Sitasi":

        gaya = st.selectbox(
            "Pilih Gaya Sitasi",
            [
                "Chicago Notes and Bibliography",
                "APA Edisi 7",
            ]
        )

        st.session_state["gaya_sitasi"] = gaya

        st.info(
            "Pilihan gaya telah disiapkan. "
            "Pembuatan sitasi otomatis belum aktif."
        )

    elif submenu == "Perbaikan Kutipan":

        st.write(
            "Fitur ini akan menyusun usulan perbaikan "
            "kutipan berdasarkan metadata sumber "
            "Aswaja An-Nahdliyah yang terverifikasi."
        )

        st.warning(
            "Mesin perbaikan kutipan belum aktif."
        )

    elif submenu == "Perbaikan Catatan Kaki":

        st.write(
            "Fitur ini akan menghasilkan catatan kaki "
            "Chicago berdasarkan sumber asli dan "
            "halaman yang dapat diverifikasi."
        )

        st.warning(
            "Mesin catatan kaki otomatis belum aktif."
        )

    elif submenu == "Perbaikan Daftar Pustaka":

        st.write(
            "Fitur ini akan menyusun daftar pustaka "
            "sesuai gaya sitasi dan metadata sumber "
            "Aswaja An-Nahdliyah yang terverifikasi."
        )

        st.warning(
            "Mesin daftar pustaka otomatis belum aktif."
        )

    elif submenu == "Pratinjau Hasil":

        st.info(
            "Pratinjau koreksi sitasi akan tersedia "
            "setelah mesin validasi dan sitasi otomatis "
            "selesai dikembangkan."
        )

    st.caption(
        "Aplikasi tidak akan mengarang judul, penulis, "
        "tahun, DOI, atau nomor halaman. Perubahan "
        "naskah memerlukan persetujuan pengguna."
    )


# ============================================================
# 13. LAPORAN VALIDASI
# ============================================================

elif menu == "📊 Laporan Validasi":

    st.subheader(submenu)

    if submenu == "Ringkasan Hasil":

        st.write(
            "**Dokumen:**",
            st.session_state.nama_dokumen or "-"
        )

        st.write(
            "**Kutipan unik terdeteksi:**",
            len(st.session_state.kutipan)
        )

        st.write(
            "**Baris daftar pustaka terdeteksi:**",
            len(st.session_state.daftar_pustaka)
        )

        st.warning(
            "Belum ada hasil validasi sumber asli. "
            "Angka di atas hanya hasil deteksi awal."
        )

    elif submenu == "Temuan Kesalahan":

        st.info(
            "Temuan kesalahan referensi akan ditampilkan "
            "setelah mesin validasi otomatis tersedia."
        )

    elif submenu == "Rekomendasi Perbaikan":

        st.info(
            "Rekomendasi perbaikan referensi "
            "Aswaja An-Nahdliyah akan ditampilkan "
            "berdasarkan hasil pemeriksaan sumber."
        )

    elif submenu == "Unduh Laporan":

        if st.session_state.hasil_baca:

            laporan = buat_laporan()

            st.download_button(
                "📥 Unduh Laporan Deteksi Awal",
                data=laporan.encode("utf-8-sig"),
                file_name="laporan_deteksi_awal_aswaja_ref_ai.txt",
                mime="text/plain"
            )

            st.caption(
                "Ini laporan deteksi awal, "
                "bukan laporan validasi referensi."
            )

        else:
            st.info(
                "Unggah dan baca dokumen terlebih dahulu."
            )

    elif submenu == "Dokumentasi Pengembangan R&D":

        st.markdown("""
        ### Dokumentasi Pengembangan ASWAJA REF-AI

        **Model penelitian:** Research and Development
        (R&D) dengan model ADDIE.

        **Tahap Development:**

        - Penyusunan struktur aplikasi Python.
        - Pembuatan repository GitHub.
        - Konfigurasi Streamlit Community Cloud.
        - Pengembangan lima menu utama.
        - Pengembangan navigasi submenu.
        - Pengembangan pembacaan PDF/DOCX.
        - Pengembangan deteksi awal kutipan.
        - Pengembangan deteksi awal daftar pustaka.

        **Rencana pengembangan berikutnya:**

        - Integrasi Generative AI.
        - Bank Literatur Aswaja An-Nahdliyah.
        - Validasi keberadaan sumber.
        - Pemeriksaan kecocokan kutipan.
        - Verifikasi halaman berdasarkan edisi.
        - Rekomendasi referensi pengganti.
        - Sitasi otomatis.
        - Laporan validasi lengkap.
        - Validasi ahli dan uji coba mahasiswa.

        Keberhasilan setiap fungsi dicatat setelah
        pengujian nyata, bukan hanya setelah kode dibuat.
        """)


# ============================================================
# 14. FOOTER
# ============================================================

st.divider()

st.caption(
    f"ASWAJA REF-AI V{VERSI} | "
    "Penelitian R&D Model ADDIE | "
    "Validasi Literatur dan Referensi Aswaja An-Nahdliyah"
)
