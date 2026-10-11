import re
import json
import difflib
import urllib.request
import urllib.parse
from io import BytesIO
from datetime import datetime

import streamlit as st
from docx import Document
from pypdf import PdfReader

st.set_page_config(page_title='ASWAJA REF-AI', page_icon='📚', layout='wide')
VERSI = '0.5.1'
JUDUL = ('Asisten Akademik Berbasis Generative AI dengan Sistem Validasi Referensi '
         'Aswaja An-Nahdliyah dan Sitasi Otomatis Karya Ilmiah Mahasiswa')

DEFAULTS = {
    'isi_dokumen': '', 'nama_dokumen': '', 'identitas': {}, 'kutipan': [],
    'daftar_pustaka': [], 'hasil_baca': False, 'hasil_pencarian': [],
    'riwayat_validasi': [], 'gaya_sitasi': 'Chicago Notes and Bibliography',
    'bank_referensi': [], 'sumber_aktif': None, 'hasil_perbandingan': None,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

MENU = {
    '🏠 Beranda': ['Dashboard', 'Panduan Penggunaan', 'Informasi Aplikasi'],
    '📤 Unggah Karya Ilmiah': ['Unggah Dokumen', 'Identitas Dokumen', 'Hasil Pembacaan', 'Deteksi Kutipan', 'Deteksi Daftar Pustaka'],
    '🔎 Validasi Literatur dan Referensi Aswaja An-Nahdliyah': [
        'Validasi Keaslian Literatur', 'Cari Sumber Asli Otomatis',
        'Bank Referensi Aswaja An-Nahdliyah', 'Unggah Sumber Asli',
        'Validasi Kesesuaian Kutipan', 'Pemeriksaan Halaman Kutipan',
        'Validasi Bibliografi', 'Pencarian Referensi Pengganti', 'Hasil Validasi'],
    '📝 Sitasi Otomatis': ['Pilihan Gaya Sitasi', 'Perbaikan Kutipan', 'Perbaikan Catatan Kaki', 'Perbaikan Daftar Pustaka', 'Pratinjau Hasil'],
    '📊 Laporan Validasi': ['Ringkasan Hasil', 'Temuan Kesalahan', 'Rekomendasi Perbaikan', 'Unduh Laporan', 'Dokumentasi Pengembangan R&D'],
}

def baca_file(file):
    data = BytesIO(file.getvalue())
    if file.name.lower().endswith('.docx'):
        doc = Document(data)
        bagian = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                bagian.append(' | '.join(c.text.strip() for c in row.cells))
        return '\n'.join(bagian), []
    if file.name.lower().endswith('.pdf'):
        reader = PdfReader(data)
        halaman = [(p.extract_text() or '') for p in reader.pages]
        return '\n'.join(halaman), halaman
    raise ValueError('Hanya PDF atau DOCX yang didukung.')

def deteksi_kutipan(teks):
    pola = r'\([^()\n]{2,90}?,\s*(?:19|20)\d{2}(?:\s*[:,]\s*[^()\n]{1,25})?\)'
    return list(dict.fromkeys(re.findall(pola, teks)))

def deteksi_pustaka(teks):
    m = list(re.finditer(r'(?im)^\s*(?:daftar pustaka|bibliografi|references)\s*$', teks))
    if not m:
        return []
    bagian = teks[m[-1].end():]
    stop = re.search(r'(?im)^\s*(?:lampiran|appendix)\s*$', bagian)
    if stop:
        bagian = bagian[:stop.start()]
    return [b.strip() for b in bagian.splitlines() if b.strip()][:500]

def normal(teks):
    return ' '.join(re.sub(r'[^\w\s]', ' ', teks.casefold()).split())

def skor(a, b):
    return round(difflib.SequenceMatcher(None, normal(a), normal(b)).ratio() * 100, 1)

@st.cache_data(ttl=3600, show_spinner=False)
def cari_katalog(query, sumber):
    if sumber == 'Crossref':
        url = 'https://api.crossref.org/works?' + urllib.parse.urlencode({'query.title': query, 'rows': 8})
    else:
        url = 'https://openlibrary.org/search.json?' + urllib.parse.urlencode({'q': query, 'limit': 8})
    req = urllib.request.Request(url, headers={'User-Agent': 'ASWAJA-REF-AI/0.5.1 (research prototype)', 'Accept': 'application/json'})
    with urllib.request.urlopen(req, timeout=18) as response:
        data = json.load(response)
    hasil = []
    if sumber == 'Crossref':
        for item in data.get('message', {}).get('items', []):
            title = (item.get('title') or [''])[0]
            if not title:
                continue
            authors = ', '.join(' '.join(filter(None, [p.get('given'), p.get('family')])) for p in item.get('author', []))
            parts = (item.get('published', {}) or {}).get('date-parts', [])
            year = str(parts[0][0]) if parts and parts[0] else '-'
            hasil.append({'judul': title, 'penulis': authors, 'tahun': year, 'penerbit': item.get('publisher', ''), 'doi': item.get('DOI', ''), 'url': item.get('URL', ''), 'sumber': sumber, 'kemiripan': skor(query, title)})
    else:
        for item in data.get('docs', []):
            title = item.get('title', '')
            if not title:
                continue
            key = item.get('key', '')
            hasil.append({'judul': title, 'penulis': ', '.join(item.get('author_name', [])[:5]), 'tahun': str(item.get('first_publish_year', '-')), 'penerbit': ', '.join(item.get('publisher', [])[:2]), 'doi': '', 'url': 'https://openlibrary.org' + key if key.startswith('/') else '', 'sumber': sumber, 'kemiripan': skor(query, title)})
    return sorted(hasil, key=lambda r: r['kemiripan'], reverse=True)

@st.cache_data(ttl=3600, show_spinner=False)
def cari_berdasarkan_jenis(query, jenis):
    """Pisahkan kategori dan jangan samakan artikel tentang kitab dengan kitabnya."""
    if jenis == 'Jurnal':
        url = 'https://api.crossref.org/works?' + urllib.parse.urlencode({
            'query.title': query, 'filter': 'type:journal-article', 'rows': 30
        })
        req = urllib.request.Request(url, headers={'User-Agent': 'ASWAJA-REF-AI/0.5.1 (research prototype)'})
        with urllib.request.urlopen(req, timeout=18) as response:
            data = json.load(response)
        hasil = []
        for item in data.get('message', {}).get('items', []):
            if item.get('type') != 'journal-article':
                continue
            title = (item.get('title') or [''])[0]
            if not title:
                continue
            authors = ', '.join(' '.join(filter(None, [p.get('given'), p.get('family')])) for p in item.get('author', []))
            parts = (item.get('published', {}) or {}).get('date-parts', [])
            hasil.append({'judul': title, 'penulis': authors,
                'tahun': str(parts[0][0]) if parts and parts[0] else '-',
                'penerbit': item.get('publisher', ''), 'doi': item.get('DOI', ''),
                'url': item.get('URL', ''), 'sumber': 'Crossref (artikel jurnal)',
                'jenis': jenis, 'kemiripan': skor(query, title)})
        return sorted(hasil, key=lambda x: x['kemiripan'], reverse=True)[:12]

    # Open Library mengindeks rekaman buku/edisi. Untuk kitab klasik dan tafsir,
    # pengarang serta edisi tetap harus dikonfirmasi secara manual.
    params = {'title': query, 'limit': 40, 'fields': 'key,title,author_name,first_publish_year,publisher,subject,edition_count,language'}
    url = 'https://openlibrary.org/search.json?' + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={'User-Agent': 'ASWAJA-REF-AI/0.5.1 (research prototype)'})
    with urllib.request.urlopen(req, timeout=18) as response:
        data = json.load(response)
    hasil = []
    for item in data.get('docs', []):
        title = item.get('title', '')
        if not title:
            continue
        # Pencarian title membatasi judul karya, bukan artikel yang membahasnya.
        # Rekaman yang tidak memiliki kemiripan judul memadai tidak ditampilkan.
        nilai = skor(query, title)
        qwords = set(normal(query).split())
        twords = set(normal(title).split())
        cakupan = len(qwords & twords) / max(len(qwords), 1)
        if nilai < 38 and cakupan < 0.65:
            continue
        subjects = ' '.join(item.get('subject', [])[:20]).casefold()
        if jenis == 'Tafsir':
            if not any(k in normal(title + ' ' + subjects) for k in ('tafsir', 'quran', 'qur an', 'koran', 'exegesis', 'commentary')):
                continue
        if jenis == 'Karya Ilmiah':
            if not any(k in normal(title + ' ' + subjects) for k in ('thesis', 'dissertation', 'skripsi', 'tesis', 'disertasi', 'academic', 'research')):
                continue
        key = item.get('key', '')
        hasil.append({'judul': title, 'penulis': ', '.join(item.get('author_name', [])[:5]),
            'tahun': str(item.get('first_publish_year', '-')),
            'penerbit': ', '.join(item.get('publisher', [])[:2]), 'doi': '',
            'url': 'https://openlibrary.org' + key if key.startswith('/') else '',
            'sumber': 'Open Library (rekaman karya)', 'jenis': jenis,
            'kemiripan': nilai, 'bahasa_katalog': item.get('language', [])})
    return sorted(hasil, key=lambda x: x['kemiripan'], reverse=True)[:12]

def klasifikasi_edisi(item):
    """Indikasi metadata, bukan putusan otentisitas atau klasifikasi filologis."""
    judul = normal(item.get('judul', ''))
    bahasa = item.get('bahasa_katalog', []) or []
    if isinstance(bahasa, str):
        bahasa = [bahasa]
    bahasa = [str(x).lower() for x in bahasa]
    if item.get('jenis') == 'Jurnal':
        return 'Artikel jurnal', 'Tidak berlaku', 'Tidak berlaku'
    if any(x in judul for x in ('commentary', 'syarah', 'sharh', 'study of', 'analisis', 'kajian terhadap')):
        kategori = 'Kemungkinan kajian atau syarah'
    elif any(x in judul for x in ('translation', 'translated', 'terjemah')):
        kategori = 'Kemungkinan terjemahan'
    elif 'eng' in bahasa or 'en' in bahasa or 'ind' in bahasa:
        kategori = 'Kemungkinan edisi terjemahan atau adaptasi'
    elif 'ara' in bahasa or 'ar' in bahasa:
        kategori = 'Kemungkinan edisi bahasa Arab'
    else:
        kategori = 'Jenis edisi belum dapat ditentukan'
    nama_bahasa = {'ara': 'Arab', 'ar': 'Arab', 'eng': 'Inggris', 'en': 'Inggris', 'ind': 'Indonesia', 'id': 'Indonesia'}
    tampilan_bahasa = ', '.join(dict.fromkeys(nama_bahasa.get(x, x) for x in bahasa)) if bahasa else 'Belum tercantum'
    return kategori, tampilan_bahasa, 'Belum terverifikasi dari katalog'

def simpan_kandidat(item):
    key = (item.get('sumber'), item.get('doi') or item.get('url') or item.get('judul'))
    if not any((r.get('sumber'), r.get('doi') or r.get('url') or r.get('judul')) == key for r in st.session_state.riwayat_validasi):
        st.session_state.riwayat_validasi.append(item.copy())
        return True
    return False

def tampil_kandidat(items, prefix):
    if not items:
        st.info('Tidak ada kandidat pada katalog ini. Ini bukan bukti bahwa referensi tidak ada.')
        return
    for i, item in enumerate(items):
        with st.container(border=True):
            st.write(f"**{i+1}. {item['judul']}**")
            st.write('Penulis:', item['penulis'] or '-')
            st.write('Tahun:', item['tahun'], '| Penerbit:', item['penerbit'] or '-')
            st.write('Katalog:', item['sumber'], '| Kemiripan judul:', str(item['kemiripan']) + '%')
            kategori_edisi, bahasa_edisi, status_edisi = klasifikasi_edisi(item)
            st.write('Jenis edisi (indikasi):', kategori_edisi)
            st.write('Bahasa (metadata katalog):', bahasa_edisi)
            st.write('Status keaslian teks:', status_edisi)
            st.caption('Pengarang asli, penerjemah, cetakan, jilid, dan halaman perlu diperiksa langsung pada edisi yang digunakan.')
            if item.get('doi'):
                st.write('DOI:', item['doi'])
            if item.get('url', '').startswith('https://'):
                st.markdown(f"[Buka rekaman katalog]({item['url']})")
            if st.button('💾 Simpan Kandidat', key=f'{prefix}_{i}'):
                if simpan_kandidat(item):
                    st.success('Metadata kandidat disimpan dalam sesi ini.')
                else:
                    st.info('Kandidat sudah tersimpan.')
    st.caption('Katalog hanya menyediakan metadata. Teks lengkap, relevansi Aswaja An-Nahdliyah, kutipan, dan halaman belum otomatis terbukti.')

def simpan_sumber(file, judul, edisi):
    teks, halaman = baca_file(file)
    if not teks.strip():
        raise ValueError('Tidak ditemukan teks yang dapat dibaca. PDF hasil scan mungkin memerlukan OCR.')
    sumber = {'id': f"{file.name}:{len(file.getvalue())}:{judul}", 'judul': judul or file.name,
              'nama_file': file.name, 'edisi': edisi, 'teks': teks, 'halaman': halaman,
              'jenis': 'PDF' if file.name.lower().endswith('.pdf') else 'DOCX'}
    if not any(s['id'] == sumber['id'] for s in st.session_state.bank_referensi):
        st.session_state.bank_referensi.append(sumber)
    st.session_state.sumber_aktif = sumber['id']
    st.session_state.hasil_perbandingan = None
    return sumber

def pilih_sumber():
    bank = st.session_state.bank_referensi
    if not bank:
        st.info('Bank referensi masih kosong. Unggah sumber asli pada submenu Unggah Sumber Asli.')
        return None
    ids = [s['id'] for s in bank]
    default = ids.index(st.session_state.sumber_aktif) if st.session_state.sumber_aktif in ids else 0
    selected = st.selectbox('Pilih sumber yang sudah diunggah', ids, index=default, format_func=lambda sid: next(s['judul'] + ' — ' + s['nama_file'] for s in bank if s['id'] == sid))
    st.session_state.sumber_aktif = selected
    return next(s for s in bank if s['id'] == selected)

def periksa_kutipan(kutipan, sumber):
    target = normal(kutipan)
    if len(target) < 12:
        return {'status': 'Kutipan terlalu pendek', 'skor': 0, 'cuplikan': '', 'lokasi': '-'}
    halaman = sumber['halaman'] if sumber['halaman'] else [sumber['teks']]
    terbaik = {'status': 'Belum ditemukan kecocokan kuat', 'skor': 0, 'cuplikan': '', 'lokasi': '-'}
    for index, teks in enumerate(halaman):
        potongan = re.split(r'(?<=[.!?؟。])\s+|\n+', teks)
        if not potongan:
            potongan = [teks]
        for bagian in potongan:
            if not bagian.strip():
                continue
            normalized = normal(bagian)
            if target in normalized:
                nilai = 100.0
            else:
                nilai = skor(kutipan, bagian[:max(300, len(kutipan) * 3)])
            if nilai > terbaik['skor']:
                terbaik = {'status': 'Kecocokan teks ditemukan' if nilai == 100 else 'Kandidat kemiripan teks',
                           'skor': nilai, 'cuplikan': bagian[:850],
                           'lokasi': f'Halaman PDF ke-{index + 1}' if sumber['halaman'] else 'Dokumen DOCX (nomor halaman tidak tersedia)'}
    if terbaik['skor'] < 60:
        terbaik['status'] = 'Belum ditemukan kecocokan kuat'
    return terbaik

def laporan():
    lines = ['ASWAJA REF-AI — LAPORAN PEMERIKSAAN AWAL', f'Versi {VERSI}', datetime.now().strftime('%d-%m-%Y'),
             'Dokumen: ' + (st.session_state.nama_dokumen or '-'),
             'Kutipan unik: ' + str(len(st.session_state.kutipan)),
             'Baris daftar pustaka: ' + str(len(st.session_state.daftar_pustaka)),
             'Metadata katalog tersimpan: ' + str(len(st.session_state.riwayat_validasi)),
             'Sumber asli diunggah dalam sesi: ' + str(len(st.session_state.bank_referensi)), '']
    for r in st.session_state.riwayat_validasi:
        lines.append('KANDIDAT KATALOG: ' + r['judul'] + ' | ' + r['sumber'] + ' | ' + r.get('url', ''))
    hasil = st.session_state.hasil_perbandingan
    if hasil:
        lines.extend(['', 'PEMERIKSAAN KUTIPAN', 'Sumber: ' + hasil['sumber'], 'Kutipan: ' + hasil['kutipan'],
                      'Status: ' + hasil['status'], 'Skor kemiripan: ' + str(hasil['skor']) + '%',
                      'Lokasi: ' + hasil['lokasi'], 'Cuplikan: ' + hasil['cuplikan']])
    lines.extend(['', 'BATASAN:', 'Kecocokan teks bukan pembuktian makna atau otentisitas kitab.',
                  'Halaman PDF tidak otomatis sama dengan halaman cetakan.',
                  'Tidak ada verifikasi otomatis tahun/edisi atau nomor halaman cetakan.',
                  'Data tersimpan hanya selama sesi Streamlit aktif.',
                  'Generative AI dan database permanen belum terintegrasi.'])
    return '\n'.join(lines)

st.title('📚 ASWAJA REF-AI')
st.caption(JUDUL)
st.info('Versi 0.5.1 | Pencarian terpisah dan pemeriksaan edisi awal. Generative AI belum terhubung.')
with st.sidebar:
    st.header('📚 ASWAJA REF-AI')
    menu = st.radio('Menu Utama', list(MENU), key='menu_utama')
    st.divider()
    st.subheader('Submenu')
    submenu = st.radio('Pilih Submenu', MENU[menu], key=f'submenu_{menu}')
    st.divider()
    st.caption('R&D ADDIE | Literatur dan Referensi Aswaja An-Nahdliyah')

if menu == '🏠 Beranda':
    if submenu == 'Dashboard':
        st.subheader('Dashboard ASWAJA REF-AI')
        cols = st.columns(4)
        for c, (label, value) in zip(cols, [('Dokumen Aktif', int(st.session_state.hasil_baca)),
                                            ('Kutipan Unik', len(st.session_state.kutipan)),
                                            ('Baris Daftar Pustaka', len(st.session_state.daftar_pustaka)),
                                            ('Sumber Diunggah', len(st.session_state.bank_referensi))]):
            c.metric(label, value)
        st.write('Validasi Literatur dan Referensi Aswaja An-Nahdliyah berbasis pemeriksaan sumber.')
    elif submenu == 'Panduan Penggunaan':
        st.subheader('Panduan Penggunaan')
        st.markdown('1. Unggah karya ilmiah PDF/DOCX.\n2. Cari metadata sumber melalui katalog.\n3. Unggah sumber asli jika teks lengkap belum tersedia.\n4. Bandingkan kutipan dengan sumber.\n5. Unduh laporan pemeriksaan awal.')
    else:
        st.subheader('Informasi Aplikasi')
        st.write('**Nama:** ASWAJA REF-AI')
        st.write('**Versi:**', VERSI)
        st.write('**Judul:**', JUDUL)
        st.write('**Model penelitian:** R&D ADDIE')
        st.write('**Penyimpanan sumber:** sementara selama sesi aktif')

elif menu == '📤 Unggah Karya Ilmiah':
    if submenu == 'Unggah Dokumen':
        st.subheader('Unggah Karya Ilmiah')
        file = st.file_uploader('Pilih PDF atau DOCX', type=['pdf', 'docx'], key='karya')
        if file and st.button('📖 Baca dan Analisis Awal', type='primary'):
            try:
                teks, _ = baca_file(file)
                if not teks.strip():
                    st.error('Teks tidak terbaca; PDF hasil scan mungkin memerlukan OCR.')
                else:
                    st.session_state.isi_dokumen = teks
                    st.session_state.nama_dokumen = file.name
                    st.session_state.kutipan = deteksi_kutipan(teks)
                    st.session_state.daftar_pustaka = deteksi_pustaka(teks)
                    st.session_state.hasil_baca = True
                    st.session_state.hasil_perbandingan = None
                    st.success('Dokumen berhasil dibaca.')
            except Exception as e:
                st.error(f'Gagal membaca dokumen: {e}')
        if st.session_state.hasil_baca:
            st.write('**Dokumen aktif:**', st.session_state.nama_dokumen)
    elif submenu == 'Identitas Dokumen':
        st.subheader('Identitas Dokumen')
        data = st.session_state.identitas
        with st.form('identitas'):
            nama = st.text_input('Nama Mahasiswa', value=data.get('nama', ''))
            nim = st.text_input('NIM', value=data.get('nim', ''))
            kampus = st.text_input('Perguruan Tinggi', value=data.get('kampus', ''))
            jenjang = st.selectbox('Jenjang', ['S1', 'S2', 'S3'], index=['S1', 'S2', 'S3'].index(data.get('jenjang', 'S2')))
            judul = st.text_area('Judul Karya Ilmiah', value=data.get('judul', ''))
            simpan = st.form_submit_button('💾 Simpan Identitas')
        if simpan:
            st.session_state.identitas = {'nama': nama, 'nim': nim, 'kampus': kampus, 'jenjang': jenjang, 'judul': judul}
            st.success('Identitas disimpan pada sesi ini.')
    elif submenu == 'Hasil Pembacaan':
        st.subheader('Hasil Pembacaan')
        if st.session_state.hasil_baca:
            st.metric('Jumlah Kata', len(st.session_state.isi_dokumen.split()))
            with st.expander('Lihat teks dokumen'):
                st.text_area('Teks', st.session_state.isi_dokumen, height=350, disabled=True)
        else:
            st.info('Unggah dokumen terlebih dahulu.')
    elif submenu == 'Deteksi Kutipan':
        st.subheader('Deteksi Kutipan')
        if not st.session_state.hasil_baca:
            st.info('Unggah dokumen terlebih dahulu.')
        else:
            st.metric('Kutipan Unik', len(st.session_state.kutipan))
            for i, k in enumerate(st.session_state.kutipan, 1):
                st.write(f'{i}. {k}')
            st.caption('Deteksi masih terbatas pada pola kutipan penulis-tahun.')
    else:
        st.subheader('Deteksi Daftar Pustaka')
        if not st.session_state.hasil_baca:
            st.info('Unggah dokumen terlebih dahulu.')
        else:
            for i, p in enumerate(st.session_state.daftar_pustaka, 1):
                st.write(f'{i}. {p}')
            st.caption('Hasil berupa baris teks, bukan entri bibliografi terstruktur.')

elif menu.startswith('🔎 Validasi Literatur'):
    st.subheader(submenu)
    if submenu in ['Validasi Keaslian Literatur', 'Cari Sumber Asli Otomatis', 'Pencarian Referensi Pengganti']:
        st.write('Pencarian metadata dipisahkan menurut jenis literatur. Hasil katalog bukan bukti keaslian teks.')
        jenis = st.selectbox('Jenis sumber', ['Kitab Klasik', 'Buku', 'Jurnal', 'Tafsir', 'Karya Ilmiah'], key=f'jenis_{submenu}')
        katalog = 'Crossref' if jenis == 'Jurnal' else 'Open Library'
        st.caption(f'Katalog utama: {katalog}. Filter jenis diterapkan sebelum hasil ditampilkan.')
        query = st.text_input('Judul atau kata kunci literatur Aswaja An-Nahdliyah', key=f'query_{submenu}')
        if st.button('🔍 Cari Referensi', key=f'cari_{submenu}', type='primary'):
            if len(query.strip()) < 4:
                st.warning('Masukkan minimal 4 karakter.')
            else:
                try:
                    with st.spinner('Mencari sumber sesuai kategori...'):
                        hasil = cari_berdasarkan_jenis(query.strip(), jenis)
                    st.session_state.hasil_pencarian = hasil
                    st.session_state['query_terakhir'] = (query.strip(), jenis, submenu)
                except Exception as e:
                    st.session_state.hasil_pencarian = []
                    st.error(f'Katalog belum dapat dihubungi: {e}')
        if st.session_state.get('query_terakhir', (None, None, None)) == (query.strip(), jenis, submenu):
            tampil_kandidat(st.session_state.hasil_pencarian, 'kandidat_' + submenu)
        st.warning('Pemeriksaan edisi ini masih berbasis metadata. Keterangan bahasa tidak selalu tersedia dan belum membuktikan kitab asli atau terjemahan.')
    elif submenu == 'Bank Referensi Aswaja An-Nahdliyah':
        st.write('Koleksi sumber asli yang sudah diunggah pada sesi Streamlit ini.')
        bank = st.session_state.bank_referensi
        st.metric('Sumber dalam sesi', len(bank))
        if not bank:
            st.info('Belum ada sumber. Buka submenu Unggah Sumber Asli.')
        for i, sumber in enumerate(bank):
            with st.expander(sumber['judul']):
                st.write('File:', sumber['nama_file'])
                st.write('Edisi:', sumber['edisi'] or 'Belum dicatat')
                st.write('Jenis:', sumber['jenis'])
                st.write('Halaman PDF terbaca:', len(sumber['halaman']) if sumber['halaman'] else 'Tidak tersedia')
                if st.button('Gunakan untuk pemeriksaan', key=f'pilihbank_{i}'):
                    st.session_state.sumber_aktif = sumber['id']
                    st.success('Sumber dipilih.')
        st.warning('Bank ini belum permanen. Dokumen dapat hilang ketika sesi berakhir atau aplikasi dimulai ulang.')
    elif submenu == 'Unggah Sumber Asli':
        st.write('Unggah PDF/DOCX kitab, buku, atau artikel Aswaja An-Nahdliyah yang Anda berhak gunakan.')
        file = st.file_uploader('Pilih sumber asli', type=['pdf', 'docx'], key='sumber_asli')
        judul = st.text_input('Judul kitab/buku/artikel')
        edisi = st.text_input('Penerbit, tahun, dan edisi/cetakan (jika diketahui)')
        if st.button('📚 Masukkan ke Bank Referensi', type='primary'):
            if not file:
                st.warning('Pilih file terlebih dahulu.')
            else:
                try:
                    sumber = simpan_sumber(file, judul.strip(), edisi.strip())
                    st.success('Sumber terbaca dan disimpan untuk sesi ini: ' + sumber['judul'])
                except Exception as e:
                    st.error(f'Gagal memproses sumber: {e}')
        st.caption('File tidak disimpan ke database permanen. PDF scan tanpa lapisan teks belum didukung.')
    elif submenu == 'Validasi Kesesuaian Kutipan':
        sumber = pilih_sumber()
        if sumber:
            kutipan = st.text_area('Tempel kutipan persis dari karya ilmiah (tanpa format sitasi)', height=130)
            if st.button('🔎 Bandingkan dengan Sumber Asli', type='primary'):
                if len(kutipan.strip()) < 12:
                    st.warning('Masukkan kutipan minimal 12 karakter.')
                else:
                    hasil = periksa_kutipan(kutipan.strip(), sumber)
                    st.session_state.hasil_perbandingan = {'sumber': sumber['judul'], 'kutipan': kutipan.strip(), **hasil}
            hasil = st.session_state.hasil_perbandingan
            if hasil and hasil['sumber'] == sumber['judul']:
                st.write('**Status:**', hasil['status'])
                st.write('**Kemiripan teks:**', str(hasil['skor']) + '%')
                st.write('**Lokasi:**', hasil['lokasi'])
                st.text_area('Cuplikan sumber untuk diperiksa', hasil['cuplikan'], height=170, disabled=True)
                st.warning('Kemiripan teks adalah petunjuk awal, bukan validasi akademik otomatis. Kutipan parafrasa memerlukan pemeriksaan makna dan konteks.')
    elif submenu == 'Pemeriksaan Halaman Kutipan':
        sumber = pilih_sumber()
        if sumber:
            if sumber['halaman']:
                st.write('Jumlah halaman PDF:', len(sumber['halaman']))
                nomor = st.number_input('Halaman PDF ke-', min_value=1, max_value=len(sumber['halaman']), value=1)
                st.text_area('Teks halaman PDF', sumber['halaman'][nomor - 1][:12000], height=300, disabled=True)
                st.warning('Halaman PDF ke-' + str(nomor) + ' belum tentu halaman cetakan ke-' + str(nomor) + '. Periksa nomor cetakan dan edisi pada sumber.')
            else:
                st.info('DOCX tidak menyediakan penomoran halaman tetap. Unggah PDF edisi yang digunakan jika ingin memeriksa halaman.')
    elif submenu == 'Validasi Bibliografi':
        st.write('Metadata kandidat dari katalog:')
        if not st.session_state.riwayat_validasi:
            st.info('Belum ada kandidat yang disimpan.')
        for i, item in enumerate(st.session_state.riwayat_validasi, 1):
            st.write(f"{i}. {item['penulis']} ({item['tahun']}). {item['judul']}. {item['penerbit']}. DOI: {item['doi'] or '-'}")
        st.warning('Metadata katalog belum tentu cocok dengan edisi yang dikutip mahasiswa.')
    else:
        st.write('### Hasil Pemeriksaan')
        st.metric('Kandidat metadata', len(st.session_state.riwayat_validasi))
        hasil = st.session_state.hasil_perbandingan
        if hasil:
            st.write('**Sumber:**', hasil['sumber'])
            st.write('**Status:**', hasil['status'])
            st.write('**Kemiripan:**', str(hasil['skor']) + '%')
            st.write('**Lokasi:**', hasil['lokasi'])
            st.text_area('Cuplikan', hasil['cuplikan'], disabled=True)
        else:
            st.info('Belum ada perbandingan kutipan.')
        st.warning('Hasil belum merupakan validasi menyeluruh terhadap keaslian, tahun, halaman cetakan, atau makna tafsir.')

elif menu == '📝 Sitasi Otomatis':
    st.subheader(submenu)
    if submenu == 'Pilihan Gaya Sitasi':
        pilihan = ['Chicago Notes and Bibliography', 'APA Edisi 7']
        st.session_state.gaya_sitasi = st.selectbox('Gaya Sitasi', pilihan, index=pilihan.index(st.session_state.gaya_sitasi))
        st.info('Pilihan disimpan untuk sesi ini.')
    else:
        st.info('Fitur ini masih tahap pengembangan. Sitasi otomatis akan dibuat setelah pemeriksaan metadata edisi dan sumber asli.')
    st.caption('Aplikasi tidak membuat nomor halaman, tahun, atau kutipan tanpa bukti sumber.')

elif menu == '📊 Laporan Validasi':
    st.subheader(submenu)
    if submenu == 'Ringkasan Hasil':
        st.write('Dokumen:', st.session_state.nama_dokumen or '-')
        st.metric('Kutipan unik', len(st.session_state.kutipan))
        st.metric('Metadata kandidat', len(st.session_state.riwayat_validasi))
        st.metric('Sumber sesi', len(st.session_state.bank_referensi))
    elif submenu == 'Unduh Laporan':
        st.download_button('📥 Unduh Laporan Pemeriksaan Awal', laporan().encode('utf-8-sig'), file_name='laporan_aswaja_ref_ai_v04.txt', mime='text/plain')
    elif submenu == 'Dokumentasi Pengembangan R&D':
        st.markdown('**R&D ADDIE — Development V0.4**\n\n- Penelusuran metadata Crossref dan Open Library.\n- Unggah sumber asli PDF/DOCX.\n- Bank referensi sementara per sesi.\n- Pencocokan awal kutipan dengan teks sumber.\n- Pemeriksaan teks per halaman PDF.\n- Laporan pemeriksaan awal.\n\n**Belum tersedia:** AI/RAG, database permanen, OCR, validasi edisi/halaman cetakan otomatis, dan generator sitasi.')
    else:
        st.info('Kesalahan dan rekomendasi yang dapat dipertanggungjawabkan memerlukan pemeriksaan lebih lanjut terhadap sumber asli.')

st.divider()
st.caption('ASWAJA REF-AI V' + VERSI + ' | R&D ADDIE | Validasi Literatur dan Referensi Aswaja An-Nahdliyah')
