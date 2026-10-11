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
VERSI = '0.6.0'
JUDUL = ('Asisten Akademik Berbasis Generative AI dengan Sistem Validasi Referensi '
         'Aswaja An-Nahdliyah dan Sitasi Otomatis Karya Ilmiah Mahasiswa')

DEFAULTS = {
    'isi_dokumen': '', 'nama_dokumen': '', 'identitas': {}, 'kutipan': [],
    'daftar_pustaka': [], 'hasil_baca': False, 'hasil_pencarian': [],
    'riwayat_validasi': [], 'gaya_sitasi': 'Chicago Notes and Bibliography',
    'bank_referensi': [], 'sumber_aktif': None, 'hasil_perbandingan': None, 'kutipan_langsung': [], 'halaman_tesis': [], 'pasangan_kutipan': {}, 'jenis_dokumen': '', 'bagian_docx': [], 'judul_terdeteksi': '', 'saringan_aswaja': {}, 'rekomendasi_pilihan': [], 'rekomendasi_draf': {},
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
        'Validasi Bibliografi', 'Pencarian Referensi Pengganti', 'Rekomendasi Kutipan & Referensi Aswaja An-Nahdliyah', 'Hasil Validasi'],
    '📝 Sitasi Otomatis': ['Pilihan Gaya Sitasi', 'Perbaikan Kutipan', 'Perbaikan Catatan Kaki', 'Perbaikan Daftar Pustaka', 'Pratinjau Hasil'],
    '📊 Laporan Validasi': ['Ringkasan Hasil', 'Temuan Kesalahan', 'Rekomendasi Perbaikan', 'Unduh Laporan', 'Dokumentasi Pengembangan R&D'],
}

def baca_file(file):
    """Mempertahankan teks asli dan pemisah halaman PDF; DOCX per paragraf."""
    data = BytesIO(file.getvalue())
    if file.name.lower().endswith('.docx'):
        doc = Document(data)
        bagian = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                bagian.append(' | '.join(c.text.strip() for c in row.cells))
        st.session_state.bagian_docx = bagian
        st.session_state.jenis_dokumen = 'DOCX'
        return '\n'.join(bagian), []
    if file.name.lower().endswith('.pdf'):
        reader = PdfReader(data)
        halaman = [(p.extract_text() or '') for p in reader.pages]
        st.session_state.bagian_docx = []
        st.session_state.jenis_dokumen = 'PDF'
        return '\n'.join(halaman), halaman
    raise ValueError('Hanya PDF atau DOCX yang didukung.')


def deteksi_judul_awal(teks):
    """Usulan judul dari halaman awal; tidak menyatakan kepastian."""
    awal = teks[:5500]
    baris = [x.strip() for x in awal.splitlines() if x.strip()]
    for i, line in enumerate(baris[:45]):
        if re.fullmatch(r'(?:TESIS|SKRIPSI|DISERTASI|THESIS)', line, re.I):
            kandidat = baris[max(0, i-9):i]
            kandidat = [x for x in kandidat if len(x) >= 15 and not re.search(r'^(?:oleh|nama|nim|program studi|fakultas|universitas)', x, re.I)]
            if kandidat:
                return ' '.join(kandidat[-3:])[:300]
    return ''


def klasifikasi_sitasi(teks):
    """Pisahkan pola penulis-tahun dari metadata penerbitan dalam tanda kurung."""
    pola = r'\([^()\n]{2,140}?,\s*(?:1[5-9]\d{2}|20\d{2})(?:\s*[:,]\s*[^()\n]{1,25})?\)'
    valid, penerbitan, perlu_tinjau = [], [], []
    kota = r'jakarta|yogyakarta|jogjakarta|bandung|surabaya|semarang|malang|bogor|solo|surakarta|medan|makassar|depok|kairo|beirut|riyadh|london|new york'
    penerbit = r'pustaka|media|press|publisher|publishing|gramedia|kalam|aksara|mizan|rajawali|grafindo|republika|kalimedia|maarif|ma.rif|bintang|erlangga|kencana|gaya|arruz|ar.ruzz|cv\.|pt\.'
    for m in re.finditer(pola, teks, flags=re.I):
        isi = m.group(0)
        kepala = re.sub(r',\s*(?:1[5-9]\d{2}|20\d{2}).*$', '', isi[1:-1], flags=re.I).strip()
        if ':' in kepala or re.search(r'^(?:' + kota + r')\b', kepala, re.I) or re.search(r'\b(?:' + penerbit + r')\b', kepala, re.I):
            penerbitan.append(isi)
        elif re.fullmatch(r'[A-ZÀ-ÿ][\wÀ-ÿ.\-\s&]{1,80}', kepala) and not re.search(r'\d', kepala):
            valid.append(isi)
        else:
            perlu_tinjau.append(isi)
    unik = lambda items: list(dict.fromkeys(items))
    return {'penulis_tahun': unik(valid), 'metadata_penerbitan': unik(penerbitan), 'perlu_tinjau': unik(perlu_tinjau)}


def deteksi_kutipan(teks):
    return klasifikasi_sitasi(teks)['penulis_tahun']

def ekstrak_kutipan_langsung(teks):
    """Kutipan bertanda petik, tidak menebak isi sumber asli."""
    pola = r'[“\"]([^“”\"\n]{20,1500})[”\"]|‘([^‘’\n]{20,1500})’'
    hasil = []
    for m in re.finditer(pola, teks):
        kutipan = next((x for x in m.groups() if x), '')
        sekitar = teks[max(0, m.start()-400):min(len(teks),m.end()+500)]
        hasil.append({'teks': kutipan, 'konteks': sekitar, 'posisi': m.start()})
    return hasil[:300]

def pasangan_referensi(konteks, pustaka):
    """Usulan pasangan saja; tidak dianggap sebagai identitas terverifikasi."""
    hasil=[]
    konteks_normal=normal(konteks)
    for entri in pustaka:
        awal=entri.split('(')[0].split(',')[0].strip()
        token=[x for x in normal(awal).split() if len(x)>3]
        if token and any(x in konteks_normal.split() for x in token):
            hasil.append(entri)
    return hasil[:5]

def ekstrak_catatan_kaki_pdf(halaman):
    """Kandidat bernomor, bukan kepastian catatan kaki atau pasangan kutipan."""
    hasil = []
    for no, teks in enumerate(halaman, 1):
        baris = [x.strip() for x in teks.splitlines() if x.strip()]
        bagian = baris[max(0, int(len(baris) * .60)):]
        for i, baris_teks in enumerate(bagian):
            m = re.match(r'^(?:\[)?(\d{1,3})(?:\]|[.)]|\s{1,3})\s*(.{12,})$', baris_teks)
            if not m:
                continue
            isi = m.group(2).strip()
            if not (re.search(r'(?:19|20)\d{2}|hlm\.?|hal\.?|jilid|cet\.?|\bvol\b', isi, re.I) or ',' in isi):
                continue
            sambungan = ' '.join(bagian[i+1:i+3])
            hasil.append({'nomor': m.group(1), 'teks': (isi + ' ' + sambungan)[:650], 'halaman_pdf': no, 'asal': 'Kandidat catatan kaki PDF'})
    return hasil


def nomor_rujukan_terdekat(data, halaman):
    """Nomor rujukan eksplisit di akhir kutipan; jangan tebak dari urutan daftar."""
    pos = data.get('posisi', -1)
    teks = st.session_state.get('isi_dokumen', '')
    if pos < 0 or pos >= len(teks):
        return None
    ujung = pos + len(data.get('teks', '')) + 2
    frag = teks[ujung:ujung+24]
    m = re.match(r'^\s*(?:\[(\d{1,3})\]|[¹²³⁴⁵⁶⁷⁸⁹⁰]+|\((\d{1,3})\))', frag)
    return next((x for x in m.groups() if x), None) if m else None

def identifikasi_halaman_kutipan(kutipan, halaman):
    """Cari halaman tesis, tidak menyamakan dengan halaman cetak."""
    target = normal(kutipan)
    for i, teks in enumerate(halaman):
        if target and target in normal(teks):
            return i + 1
    awal = normal(kutipan)[:65]
    for i, teks in enumerate(halaman):
        if awal and awal in normal(teks):
            return i + 1
    return None


def kandidat_sumber_kutipan(data, halaman, pustaka):
    halaman_ke = identifikasi_halaman_kutipan(data['teks'], halaman)
    kandidat = []
    nomor = nomor_rujukan_terdekat(data, halaman)
    if halaman_ke:
        catatan = ekstrak_catatan_kaki_pdf(halaman)
        kandidat = [item for item in catatan if abs(item['halaman_pdf'] - halaman_ke) <= 1]
        kandidat.sort(key=lambda c: (0 if nomor and c['nomor'] == nomor else 1, abs(c['halaman_pdf']-halaman_ke)))
    refs = pasangan_referensi(data['konteks'], pustaka)
    return halaman_ke, kandidat[:12], refs

def lokasi_kutipan_dokumen(data):
    """Temukan lokasi kutipan tanpa menyamakan nomor cetak dan indeks PDF."""
    halaman = st.session_state.get('halaman_tesis', [])
    teks = st.session_state.get('isi_dokumen', '')
    posisi = data.get('posisi', 0)
    sebelum = teks[:posisi]
    heading = list(re.finditer(r'(?im)^\s*(?:BAB\s+(?:[IVXLC]+|\d+)\b[^\n]{0,110}|CHAPTER\s+\d+\b[^\n]{0,110})', sebelum))
    bab = heading[-1].group(0).strip() if heading else 'Belum terdeteksi pada dokumen'
    no_pdf = None
    teks_halaman = ''
    if halaman:
        akumulasi = 0
        for i, h in enumerate(halaman, 1):
            akhir = akumulasi + len(h)
            if posisi < akhir or i == len(halaman):
                no_pdf, teks_halaman = i, h
                break
            akumulasi = akhir + 1
    nomor_cetak = 'Belum dapat dipastikan'
    if teks_halaman:
        baris = [x.strip() for x in teks_halaman.splitlines() if x.strip()]
        kandidat = baris[:3] + baris[-3:]
        angka = [x for x in kandidat if re.fullmatch(r'(?:[0-9]{1,4}|[ivxlcdmIVXLCDM]{1,12})', x)]
        if len(set(angka)) == 1:
            nomor_cetak = angka[0] + ' (indikasi, periksa PDF)'
    posisi_docx = None
    if not halaman and st.session_state.get('jenis_dokumen') == 'DOCX':
        posisi_docx = sebelum.count('\n') + 1
    return bab, no_pdf, nomor_cetak, posisi_docx

def tampil_mesin_pasangan(data, nomor):
    halaman = st.session_state.get('halaman_tesis', [])
    pustaka = st.session_state.daftar_pustaka
    halaman_ke, catatan, refs = kandidat_sumber_kutipan(data, halaman, pustaka)
    st.markdown('#### Mesin Pencocokan Kutipan dan Referensi')
    st.write('**Lokasi kutipan dalam dokumen:**',
             f'Halaman PDF ke-{halaman_ke}' if halaman_ke else 'Belum ditemukan secara pasti')
    st.caption('Nomor halaman PDF tesis bukan nomor halaman sumber kitab.')
    nomor_rujukan = nomor_rujukan_terdekat(data, halaman)
    st.write('**Nomor rujukan yang melekat pada kutipan:**', nomor_rujukan or 'Tidak terdeteksi secara eksplisit')
    opsi = ['Belum dapat dipastikan']
    for c in catatan:
        opsi.append(f"Catatan kaki kandidat {c['nomor']} | PDF hlm. {c['halaman_pdf']} | {c['teks'][:130]}")
    for r in refs:
        opsi.append('Daftar pustaka kandidat | ' + r[:160])
    st.write('**Kandidat sumber yang terbaca dari tesis:**')
    if len(opsi) == 1:
        st.warning('Belum ada catatan kaki atau daftar pustaka yang dapat dipasangkan secara andal. Ini bukan bukti bahwa sumber tidak ada.')
    pilihan = st.selectbox('Periksa kandidat sumber kutipan', opsi, key=f'pasangan_{nomor}')
    if pilihan != opsi[0]:
        st.info('Kandidat belum terverifikasi. Periksa teks catatan kaki dan identitas edisinya sebelum validasi sumber asli.')
    st.session_state.pasangan_kutipan[str(nomor)] = pilihan
    return pilihan


def tampilkan_kutipan_tesis():
    if not st.session_state.hasil_baca:
        st.info('Unggah dan baca tesis terlebih dahulu.');return None
    kutipan=st.session_state.get('kutipan_langsung', [])
    st.write('**Dokumen:**',st.session_state.nama_dokumen)
    identitas = st.session_state.get('identitas', {})
    st.write('**Judul karya ilmiah:**', identitas.get('judul') or ((st.session_state.get('judul_terdeteksi') or 'Belum diisi pada Identitas Dokumen') + ' (usulan, belum dikonfirmasi)' if st.session_state.get('judul_terdeteksi') else 'Belum diisi pada Identitas Dokumen'))
    st.caption('Judul dapat diisi melalui Unggah Karya Ilmiah → Identitas Dokumen. Nama file bukan pengganti judul tesis.')
    st.metric('Kutipan langsung terdeteksi',len(kutipan))
    if not kutipan:
        st.warning('Belum terdeteksi kutipan langsung bertanda petik. Anda dapat menggunakan pemeriksaan manual yang sudah tersedia.');return None
    nomor=st.selectbox('Pilih kutipan dari tesis',range(len(kutipan)),format_func=lambda i:f'Kutipan {i+1}: {kutipan[i]["teks"][:90]}',key='pilih_kutipan_tesis')
    data=kutipan[nomor]
    bab, halaman_pdf, halaman_cetak, posisi_docx = lokasi_kutipan_dokumen(data)
    st.write('**BAB / bagian karya ilmiah:**', bab)
    st.write('**Halaman PDF dokumen:**', f'Ke-{halaman_pdf}' if halaman_pdf else ('Tidak berlaku untuk DOCX' if st.session_state.get('jenis_dokumen') == 'DOCX' else 'Belum ditemukan'))
    if posisi_docx:
        st.write('**Lokasi DOCX:**', f'Paragraf teks ke-{posisi_docx} (bukan nomor halaman cetak)')
    st.write('**Halaman cetak karya ilmiah:**', halaman_cetak)
    st.caption('Halaman dokumen karya ilmiah berbeda dari halaman kitab atau buku yang dikutip.')
    st.text_area('Kutipan asli dalam tesis (tidak diubah)',data['teks'],height=130,disabled=True)
    st.text_area('Konteks kutipan dalam tesis',data['konteks'],height=150,disabled=True)
    tampil_mesin_pasangan(data, nomor)
    pasangan=pasangan_referensi(data['konteks'],st.session_state.daftar_pustaka)
    if pasangan:
        st.write('**Kemungkinan sumber yang tercantum dalam daftar pustaka:**')
        for ref in pasangan:st.write('- '+ref)
        st.caption('Pasangan referensi masih dugaan berdasarkan nama. Periksa catatan kaki dan daftar pustaka sebelum validasi.')
    else:
        st.warning('Sumber kutipan belum dapat dipasangkan dengan daftar pustaka secara pasti. Catatan kaki DOCX belum diekstrak secara khusus.')
    return data

def deteksi_pustaka(teks):
    m = list(re.finditer(r'(?im)^\s*(?:daftar pustaka|bibliografi|references)\s*$', teks))
    if not m:
        return []
    bagian = teks[m[-1].end():]
    stop = re.search(r'(?im)^\s*(?:lampiran|appendix)\s*$', bagian)
    if stop:
        bagian = bagian[:stop.start()]
    return [b.strip() for b in bagian.splitlines() if b.strip()][:500]

def saring_referensi_aswaja(pustaka, halaman, jenis_dokumen):
    """Penyaringan awal konservatif; bukan penetapan otoritas kitab secara otomatis."""
    catatan = ekstrak_catatan_kaki_pdf(halaman) if halaman else []
    daftar = ([{'teks': x, 'asal': 'Daftar pustaka'} for x in pustaka] +
             [{'teks': x['teks'], 'asal': f"Catatan kaki kandidat {x['nomor']} (PDF {x['halaman_pdf']})"} for x in catatan])
    penanda = [
        r'aswaja', r'an[- ]nahdliy', r'nahdlatul ulama', r'\bnu\b',
        r'al[- ]ghazal', r'al[- ]nawaw', r'asy[- ]syafi', r'al[- ]syafi',
        r'ibn hajar', r'ibnu hajar', r'ibn katsir', r'ibnu katsir',
        r'ihya[’\' ]+ulum', r'إحياء علوم الدين', r'فتح المعين',
        r'fathul mu[’\']in', r'فتح القريب', r'fathul qarib',
        r'tafsir jalalain', r'تفسير الجلالين', r'عقيدة العوام',
        r'aqidatul awam', r'سفينة النجاة', r'safinatun najah',
        r'رسالة أهل السنة', r'risalah ahlussunnah', r'النووي', r'الغزالي',
        r'الشافعي', r'الأشعري', r'الماتريدي', r'ash[’\']ari', r'maturidi',
        r'kitab kuning', r'ahlussunnah wal jama',
    ]
    cocok, tinjau = [], []
    seen = set()
    for item in daftar:
        teks = item['teks'].strip()
        if len(teks) < 9 or teks.casefold() in seen:
            continue
        seen.add(teks.casefold())
        if any(re.search(p, teks, re.I) for p in penanda):
            cocok.append(item)
        else:
            tinjau.append(item)
    if cocok:
        status = 'ditemukan'
    elif tinjau:
        status = 'perlu_tinjau'
    elif jenis_dokumen == 'DOCX':
        status = 'belum_pasti'  # Catatan kaki DOCX belum didukung.
    else:
        status = 'tidak_ditemukan'
    return {'status': status, 'cocok': cocok, 'tinjau': tinjau, 'catatan': len(catatan)}


def tampil_saringan_aswaja():
    if not st.session_state.hasil_baca:
        st.info('Unggah dan baca karya ilmiah terlebih dahulu sebelum validasi.')
        return False
    data = st.session_state.get('saringan_aswaja', {})
    status = data.get('status', 'belum_pasti')
    if status == 'ditemukan':
        st.success(f"Ditemukan {len(data.get('cocok', []))} kandidat literatur atau referensi Aswaja An-Nahdliyah. Identitas sumber perlu dikonfirmasi sebelum validasi.")
        return True
    if status == 'tidak_ditemukan':
        st.warning('Tidak ditemukan literatur atau referensi Aswaja An-Nahdliyah pada hasil ekstraksi. Proses validasi tidak dilanjutkan.')
    else:
        st.warning('Referensi Aswaja An-Nahdliyah belum dapat dipastikan dari hasil pembacaan. Periksa daftar pustaka atau catatan kaki sebelum validasi; proses otomatis belum dilanjutkan.')
        if data.get('tinjau'):
            with st.expander('Lihat referensi yang perlu diperiksa'):
                for item in data['tinjau'][:100]:
                    st.write(f"{item['asal']}: {item['teks']}")
    return False

def kandidat_narasi_rekomendasi(teks, halaman):
    """Kandidat paragraf karya ilmiah, bukan kutipan atau rekomendasi yang sudah tervalidasi."""
    hasil = []
    if halaman:
        sumber = [(i + 1, h) for i, h in enumerate(halaman)]
    else:
        sumber = [(None, teks)]
    for nomor, isi in sumber:
        for bagian in re.split(r'\n\s*\n|(?<=\.)\s*\n', isi):
            bagian = ' '.join(bagian.split())
            if not 110 <= len(bagian) <= 1800:
                continue
            if re.search(r'^(?:daftar pustaka|bibliografi|references|lampiran|abstract|abstrak)\b', bagian, re.I):
                continue
            if len(bagian.split()) < 18:
                continue
            hasil.append({'narasi': bagian, 'halaman_pdf': nomor})
            if len(hasil) >= 250:
                return hasil
    return hasil


def tampil_rekomendasi_aswaja():
    st.markdown('### Rekomendasi Kutipan & Referensi Aswaja An-Nahdliyah')
    st.caption('Submenu penyusunan usulan berdasarkan narasi asli; tidak otomatis mengubah karya ilmiah atau mengklaim validitas sumber.')
    if not st.session_state.hasil_baca:
        st.info('Unggah dan baca karya ilmiah terlebih dahulu.')
        return
    status = st.session_state.get('saringan_aswaja', {}).get('status')
    if status == 'tidak_ditemukan':
        st.info('Belum ditemukan referensi Aswaja An-Nahdliyah. Validasi referensi lama tetap dihentikan; Anda dapat menyiapkan rekomendasi baru di sini.')
    elif status in ('perlu_tinjau', 'belum_pasti'):
        st.warning('Identifikasi referensi awal belum pasti. Rekomendasi di sini bukan bukti bahwa tesis tidak memiliki sumber Aswaja An-Nahdliyah.')
    else:
        st.success('Ada kandidat referensi Aswaja An-Nahdliyah. Submenu ini juga dapat dipakai untuk usulan penguatan narasi.')
    pilihan_jumlah = st.selectbox('Jumlah rekomendasi yang ditargetkan', ['Otomatis', 5, 10, 15, 20, 30], key='target_rekomendasi_060')
    st.caption('Target bukan kuota wajib. Jumlah rekomendasi valid ditentukan oleh bukti sumber yang berhasil diperiksa.')
    kandidat = kandidat_narasi_rekomendasi(st.session_state.isi_dokumen, st.session_state.get('halaman_tesis', []))
    if not kandidat:
        st.warning('Belum ada paragraf yang cukup jelas untuk dipilih. Periksa hasil ekstraksi dokumen.')
        return
    batas = len(kandidat) if pilihan_jumlah == 'Otomatis' else min(int(pilihan_jumlah), len(kandidat))
    st.write(f'Paragraf kandidat terbaca: {len(kandidat)}. Pilih narasi yang benar-benar membutuhkan penguatan referensi.')
    nomor = st.selectbox('Pilih narasi dari karya ilmiah', range(len(kandidat)), format_func=lambda i: f"Paragraf {i+1} | " + (f"PDF hlm. {kandidat[i]['halaman_pdf']} | " if kandidat[i]['halaman_pdf'] else 'DOCX | ') + kandidat[i]['narasi'][:90], key='narasi_rekomendasi_060')
    item = kandidat[nomor]
    st.text_area('Narasi asli (tidak diubah)', item['narasi'], height=180, disabled=True)
    st.write('**Halaman karya ilmiah:**', f"PDF ke-{item['halaman_pdf']}" if item['halaman_pdf'] else 'Nomor halaman DOCX belum tersedia')
    st.markdown('#### Usulan literatur untuk narasi ini')
    judul = st.text_input('Judul kitab atau literatur Aswaja An-Nahdliyah kandidat', key=f'judul_usulan_060_{nomor}', placeholder='Isi dari sumber yang ditemukan, bukan judul buatan')
    teks_kutipan = st.text_area('Kutipan sumber asli (hanya jika sudah dibaca dari kitab/buku)', key=f'kutipan_usulan_060_{nomor}', height=110)
    footnote = st.text_area('Identitas sumber / footnote kandidat (pengarang, judul, edisi, halaman)', key=f'footnote_usulan_060_{nomor}', height=90)
    sumber = st.session_state.get('bank_referensi', [])
    if sumber:
        st.caption('Untuk menguji kesesuaian teks, pilih sumber asli yang sudah diunggah ke Bank Referensi.')
        pilih = st.selectbox('Sumber asli untuk pemeriksaan teks', range(len(sumber)), format_func=lambda i: sumber[i]['judul'], key=f'sumber_usulan_060_{nomor}')
        if st.button('Periksa kutipan pada sumber yang diunggah', key=f'cek_usulan_060_{nomor}'):
            if len(teks_kutipan.strip()) < 12:
                st.warning('Masukkan teks kutipan sumber minimal 12 karakter.')
            else:
                hasil = periksa_kutipan(teks_kutipan.strip(), sumber[pilih])
                st.session_state.rekomendasi_draf[str(nomor)] = {'narasi': item['narasi'], 'halaman_pdf': item['halaman_pdf'], 'judul': judul, 'kutipan': teks_kutipan, 'footnote': footnote, 'sumber': sumber[pilih]['judul'], 'hasil': hasil}
                st.write('**Hasil pemeriksaan awal:**', hasil.get('status', 'Belum pasti'))
                st.write('**Kemiripan teks:**', str(hasil.get('skor', 0)) + '%')
                st.write('**Lokasi dalam sumber:**', hasil.get('lokasi', 'Belum ditemukan'))
                st.warning('Pencocokan teks belum membuktikan ketepatan edisi, makna, dan halaman cetak. Lanjutkan ke Pemeriksaan Halaman Kutipan.')
    else:
        st.info('Belum ada sumber asli dalam sesi. Cari kandidat pada Pencarian Referensi Pengganti, lalu unggah PDF/DOCX sumber asli sebelum memvalidasi teks.')
    if st.button('Simpan sebagai usulan belum terverifikasi', key=f'simpan_usulan_060_{nomor}'):
        st.session_state.rekomendasi_draf[str(nomor)] = {'narasi': item['narasi'], 'halaman_pdf': item['halaman_pdf'], 'judul': judul, 'kutipan': teks_kutipan, 'footnote': footnote, 'status': 'Belum terverifikasi'}
        st.success('Usulan tersimpan untuk sesi ini. Belum dinyatakan sebagai kutipan valid.')
    st.metric('Usulan tersimpan dalam sesi', len(st.session_state.rekomendasi_draf))
    st.caption(f'Target pilihan: {pilihan_jumlah}; batas tampilan usulan: {batas}. Jangan memasukkan rekomendasi ke karya ilmiah sebelum validasi teks, halaman dan persetujuan pengguna.')
    st.markdown('**Tahap berikutnya:** Validasi Kesesuaian Kutipan → Pemeriksaan Halaman Kutipan → Hasil Validasi.')


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
    req = urllib.request.Request(url, headers={'User-Agent': 'ASWAJA-REF-AI/0.5.5 (research prototype)', 'Accept': 'application/json'})
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

PADANAN_JUDUL_ARAB = {
    'ihya ulumuddin': 'إحياء علوم الدين',
    'ihya ulum al din': 'إحياء علوم الدين',
    'ihya ulum ud din': 'إحياء علوم الدين',
    'ihya ulumiddin': 'إحياء علوم الدين',
    'ihya ulumuddin imam ghazali': 'إحياء علوم الدين',
}

RUJUKAN_KITAB = {
    'Maktabah Syamilah': 'https://shamela.ws/search',
    'Maktabah Waqfiyyah': 'https://waqfeya.net',
    'IslamWeb Library': 'https://www.islamweb.net/ar/library/',
    'Mishkat al-Islamiyyah': 'https://www.almeshkat.net',
    'Internet Archive': 'https://archive.org/search',
    'Altafsir': 'https://www.altafsir.com',
    'WorldCat': 'https://search.worldcat.org/search',
}

def judul_arab(query):
    teks = normal(query)
    if 'ihya' in teks and ('ulum' in teks or 'ul um' in teks):
        return 'إحياء علوم الدين'
    return PADANAN_JUDUL_ARAB.get(teks, query)

def bahasa_arab_terkonfirmasi(item):
    bahasa = item.get('bahasa_katalog', []) or []
    if isinstance(bahasa, str):
        bahasa = [bahasa]
    kode = {str(x).strip().lower() for x in bahasa}
    return bool(kode & {'ara', 'ar', 'arabic', 'العربية', 'عربي'})

@st.cache_data(ttl=1800, show_spinner=False)
def cari_archive(query, jumlah=12):
    # Advanced Search API menyediakan metadata, bukan jaminan kebenaran edisi.
    ekspresi = f'(title:("{query.replace(chr(34), "")}") AND mediatype:texts)'
    params = [('q', ekspresi), ('rows', str(jumlah)), ('page', '1'),
              ('output', 'json'), ('fl[]', 'identifier'), ('fl[]', 'title'),
              ('fl[]', 'creator'), ('fl[]', 'date'), ('fl[]', 'publisher'),
              ('fl[]', 'language')]
    url = 'https://archive.org/advancedsearch.php?' + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={'User-Agent':'ASWAJA-REF-AI/0.5.5 research'})
    with urllib.request.urlopen(req, timeout=20) as response:
        data = json.load(response)
    hasil=[]
    for x in data.get('response', {}).get('docs', []):
        judul=x.get('title','')
        if not isinstance(judul,str) or not judul.strip():
            continue
        def teks(v):
            return ', '.join(str(z) for z in v) if isinstance(v,list) else str(v or '')
        ident=x.get('identifier','')
        hasil.append({'judul':judul,'penulis':teks(x.get('creator')),
            'tahun':teks(x.get('date')) or '-', 'penerbit':teks(x.get('publisher')),
            'doi':'', 'url':'https://archive.org/details/'+urllib.parse.quote(ident,safe=''),
            'sumber':'Internet Archive (metadata koleksi)', 'jenis':'Kitab Klasik',
            'kemiripan':skor(query,judul), 'bahasa_katalog':x.get('language',[]),
            'status_verifikasi':'Belum diperiksa'})
    return sorted(hasil,key=lambda x:x['kemiripan'],reverse=True)

def tampil_rujukan_kitab(query):
    arab=judul_arab(query)
    st.write('**Judul pencarian Arab:**', arab)
    st.caption('Tautan di bawah membuka perpustakaan eksternal. Ketersediaan pencarian otomatis dan akses teks berbeda-beda.')
    for nama, alamat in RUJUKAN_KITAB.items():
        if nama == 'Internet Archive':
            tujuan=alamat+'?'+urllib.parse.urlencode({'query':arab})
        elif nama == 'WorldCat':
            tujuan=alamat+'?'+urllib.parse.urlencode({'q':arab})
        elif nama == 'Maktabah Syamilah':
            tujuan=alamat+'?'+urllib.parse.urlencode({'query':arab})
        else:
            tujuan=alamat
        st.markdown(f'- [{nama}]({tujuan})')
    st.info('Periksa nama pengarang, bahasa, penerbit, cetakan, jilid dan halaman langsung pada edisi yang dibuka. Jangan samakan halaman edisi berbeda.')

@st.cache_data(ttl=3600, show_spinner=False)
def cari_berdasarkan_jenis(query, jenis):
    """Pisahkan kategori dan jangan samakan artikel tentang kitab dengan kitabnya."""
    if jenis == 'Jurnal':
        url = 'https://api.crossref.org/works?' + urllib.parse.urlencode({
            'query.title': query, 'filter': 'type:journal-article', 'rows': 30
        })
        req = urllib.request.Request(url, headers={'User-Agent': 'ASWAJA-REF-AI/0.5.5 (research prototype)'})
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
    req = urllib.request.Request(url, headers={'User-Agent': 'ASWAJA-REF-AI/0.5.5 (research prototype)'})
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
             'Judul karya ilmiah: ' + (st.session_state.identitas.get('judul') or 'Belum diisi'),
             'Kutipan unik: ' + str(len(st.session_state.kutipan)),
             'Baris daftar pustaka: ' + str(len(st.session_state.daftar_pustaka)),
             'Metadata katalog tersimpan: ' + str(len(st.session_state.riwayat_validasi)),
             'Sumber asli diunggah dalam sesi: ' + str(len(st.session_state.bank_referensi)), '']
    for r in st.session_state.riwayat_validasi:
        lines.append('KANDIDAT KATALOG: ' + r['judul'] + ' | ' + r['sumber'] + ' | ' + r.get('url', ''))
    for nomor, pasangan in st.session_state.get('pasangan_kutipan', {}).items():
        indeks = int(nomor)
        data_kutipan = st.session_state.get('kutipan_langsung', [])
        if indeks < len(data_kutipan):
            bab, halaman_pdf, halaman_cetak, _ = lokasi_kutipan_dokumen(data_kutipan[indeks])
            lines.append(f'KUTIPAN {indeks+1} | BAB: {bab} | Halaman PDF karya ilmiah: {halaman_pdf or "Tidak tersedia"} | Halaman cetak: {halaman_cetak}')
        lines.append('KUTIPAN ' + str(indeks+1) + ' | Kandidat dari tesis: ' + pasangan)
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
st.info('Versi 0.5.9 | Katalog terpisah, pembanding Internet Archive, dan tautan penelusuran kitab Arab. Validasi teks tetap memerlukan sumber asli.')
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
                teks, halaman_tesis = baca_file(file)
                if not teks.strip():
                    st.error('Teks tidak terbaca; PDF hasil scan mungkin memerlukan OCR.')
                else:
                    st.session_state.halaman_tesis = halaman_tesis
                    st.session_state.judul_terdeteksi = deteksi_judul_awal(teks)
                    st.session_state.identitas = {} if file.name != st.session_state.get('nama_dokumen') else st.session_state.identitas
                    st.session_state.pasangan_kutipan = {}
                    st.session_state.rekomendasi_draf = {}
                    st.session_state.isi_dokumen = teks
                    st.session_state.nama_dokumen = file.name
                    st.session_state.kutipan = deteksi_kutipan(teks)
                    st.session_state.kutipan_langsung = ekstrak_kutipan_langsung(teks)
                    st.session_state.daftar_pustaka = deteksi_pustaka(teks)
                    st.session_state.saringan_aswaja = saring_referensi_aswaja(st.session_state.daftar_pustaka, halaman_tesis, st.session_state.jenis_dokumen)
                    st.session_state.hasil_baca = True
                    st.session_state.hasil_perbandingan = None
                    st.success('Dokumen berhasil dibaca.')
                    tampil_saringan_aswaja()
            except Exception as e:
                st.error(f'Gagal membaca dokumen: {e}')
        if st.session_state.hasil_baca:
            st.write('**Dokumen aktif:**', st.session_state.nama_dokumen)
            st.write('**Jenis dokumen:**', st.session_state.get('jenis_dokumen') or 'Belum diketahui')
            if st.session_state.get('judul_terdeteksi'):
                st.caption('Usulan judul dari halaman awal (perlu konfirmasi): ' + st.session_state.judul_terdeteksi)
    elif submenu == 'Identitas Dokumen':
        st.subheader('Identitas Dokumen')
        data = st.session_state.identitas
        with st.form('identitas'):
            nama = st.text_input('Nama Mahasiswa', value=data.get('nama', ''))
            nim = st.text_input('NIM', value=data.get('nim', ''))
            kampus = st.text_input('Perguruan Tinggi', value=data.get('kampus', ''))
            jenjang = st.selectbox('Jenjang', ['S1', 'S2', 'S3'], index=['S1', 'S2', 'S3'].index(data.get('jenjang', 'S2')))
            judul = st.text_area('Judul Karya Ilmiah', value=data.get('judul', '') or st.session_state.get('judul_terdeteksi', ''))
            st.caption('Judul hasil pembacaan otomatis hanya usulan. Periksa sebelum menyimpan.')
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
            tampilkan_kutipan_tesis()
            hasil_sitasi = klasifikasi_sitasi(st.session_state.isi_dokumen)
            st.metric('Sitasi penulis-tahun terdeteksi', len(hasil_sitasi['penulis_tahun']))
            for i, k in enumerate(hasil_sitasi['penulis_tahun'], 1):
                st.write(f'{i}. {k}')
            with st.expander(f"Metadata penerbitan, bukan sitasi penulis-tahun ({len(hasil_sitasi['metadata_penerbitan'])})"):
                for i, k in enumerate(hasil_sitasi['metadata_penerbitan'], 1):
                    st.write(f'{i}. {k}')
            with st.expander(f"Pola ambigu yang perlu diperiksa ({len(hasil_sitasi['perlu_tinjau'])})"):
                for i, k in enumerate(hasil_sitasi['perlu_tinjau'], 1):
                    st.write(f'{i}. {k}')
            st.caption('Klasifikasi berbasis pola, bukan verifikasi bibliografi. Catatan kaki harus ditautkan dengan nomor rujukan yang jelas.')
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
    if submenu == 'Rekomendasi Kutipan & Referensi Aswaja An-Nahdliyah':
        tampil_rekomendasi_aswaja()
    izin_validasi = tampil_saringan_aswaja() if submenu != 'Rekomendasi Kutipan & Referensi Aswaja An-Nahdliyah' else False
    if submenu in ['Validasi Keaslian Literatur', 'Cari Sumber Asli Otomatis', 'Pencarian Referensi Pengganti']:
        st.write('Pencarian metadata dipisahkan menurut jenis literatur. Hasil katalog bukan bukti keaslian teks.')
        jenis = st.selectbox('Jenis sumber', ['Kitab Klasik', 'Buku', 'Jurnal', 'Tafsir', 'Karya Ilmiah'], key=f'jenis_{submenu}')
        katalog = 'Crossref' if jenis == 'Jurnal' else ('Internet Archive dan katalog pembanding' if jenis in ('Kitab Klasik', 'Tafsir') else 'Open Library')
        st.caption(f'Katalog: {katalog}. Pencarian lintas situs tidak berarti semua situs menyediakan API.')
        pilihan = st.session_state.get('saringan_aswaja', {}).get('cocok', [])
        if pilihan:
            label_pilihan = [f"{item['asal']}: {item['teks'][:160]}" for item in pilihan]
            dipilih = st.selectbox('Referensi atau footnote yang terdeteksi dalam tesis', range(len(pilihan)), format_func=lambda i: label_pilihan[i], key=f'pilih_ref_{submenu}')
            sumber_pilihan = pilihan[dipilih]['teks']
            kunci = f'query_{submenu}'
            penanda_kunci = f'asal_query_{submenu}'
            if st.session_state.get(penanda_kunci) != sumber_pilihan:
                st.session_state[kunci] = sumber_pilihan[:220]
                st.session_state[penanda_kunci] = sumber_pilihan
        query = st.text_input('Judul atau kata kunci literatur Aswaja An-Nahdliyah', key=f'query_{submenu}', disabled=not izin_validasi)
        if not izin_validasi:
            st.stop()
        if st.button('🔍 Cari Referensi', key=f'cari_{submenu}', type='primary'):
            if len(query.strip()) < 4:
                st.warning('Masukkan minimal 4 karakter.')
            else:
                try:
                    with st.spinner('Mencari sumber sesuai kategori...'):
                        if jenis in ('Kitab Klasik', 'Tafsir'):
                            semua = cari_archive(judul_arab(query.strip()), jumlah=40)
                            # Kebijakan ketat: jangan tampilkan bahasa Inggris atau bahasa tidak diketahui
                            # sebagai kitab Arab. Hasil yang tidak diketahui tetap belum terverifikasi.
                            hasil = [r for r in semua if bahasa_arab_terkonfirmasi(r)]
                            st.session_state['jumlah_ditolak_bahasa'] = len(semua) - len(hasil)
                        else:
                            hasil = cari_berdasarkan_jenis(query.strip(), jenis)
                    st.session_state.hasil_pencarian = hasil
                    st.session_state['query_terakhir'] = (query.strip(), jenis, submenu)
                except Exception as e:
                    st.session_state.hasil_pencarian = []
                    st.session_state['query_terakhir'] = (query.strip(), jenis, submenu)
                    st.error(f'Katalog belum dapat dihubungi: {e}')
        if st.session_state.get('query_terakhir', (None, None, None)) == (query.strip(), jenis, submenu):
            if jenis in ('Kitab Klasik', 'Tafsir'):
                st.info('Filter ketat bahasa Arab aktif. Edisi Inggris dan entri tanpa metadata bahasa Arab tidak ditampilkan sebagai kitab Arab.')
                if st.session_state.get('jumlah_ditolak_bahasa', 0):
                    st.caption(f"{st.session_state['jumlah_ditolak_bahasa']} kandidat disisihkan karena bukan bahasa Arab terkonfirmasi.")
            tampil_kandidat(st.session_state.hasil_pencarian, 'kandidat_' + submenu)
        if jenis in ('Kitab Klasik', 'Tafsir') and query.strip():
            st.divider()
            st.write('#### Rujukan pembanding kitab Arab')
            tampil_rujukan_kitab(query.strip())
        st.warning('Hasil katalog belum membuktikan keaslian teks, relevansi Aswaja An-Nahdliyah, atau ketepatan halaman. Jika katalog gagal, gunakan tautan pembanding dan unggah sumber asli.')
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
        st.markdown('### Tahap 1. Kutipan dan sumber sebagaimana tertulis dalam tesis')
        data_kutipan = tampilkan_kutipan_tesis()
        st.markdown('### Tahap 2. Pemeriksaan terhadap sumber asli')
        sumber = pilih_sumber()
        if sumber:
            nilai_awal = data_kutipan['teks'] if data_kutipan else ''
            kutipan = st.text_area('Kutipan yang akan dibandingkan (dapat menggunakan hasil deteksi)', value=nilai_awal, height=130, key='teks_validasi_054')
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
        st.write('### Kutipan dalam tesis sebelum hasil validasi')
        tampilkan_kutipan_tesis()
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
        st.download_button('📥 Unduh Laporan Pemeriksaan Awal', laporan().encode('utf-8-sig'), file_name='laporan_aswaja_ref_ai_v055.txt', mime='text/plain')
    elif submenu == 'Dokumentasi Pengembangan R&D':
        st.markdown('**R&D ADDIE — Development V0.4**\n\n- Penelusuran metadata Crossref dan Open Library.\n- Unggah sumber asli PDF/DOCX.\n- Bank referensi sementara per sesi.\n- Pencocokan awal kutipan dengan teks sumber.\n- Pemeriksaan teks per halaman PDF.\n- Laporan pemeriksaan awal.\n\n**Belum tersedia:** AI/RAG, database permanen, OCR, validasi edisi/halaman cetakan otomatis, dan generator sitasi.')
    else:
        st.info('Kesalahan dan rekomendasi yang dapat dipertanggungjawabkan memerlukan pemeriksaan lebih lanjut terhadap sumber asli.')

st.divider()
st.caption('ASWAJA REF-AI V' + VERSI + ' | R&D ADDIE | Validasi Literatur dan Referensi Aswaja An-Nahdliyah')
