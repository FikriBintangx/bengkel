import os
import re
import time
from docx import Document
from flask import Flask, request, render_template_string, send_file, redirect, url_for, jsonify
import google.generativeai as genai
from pypdf import PdfReader, PdfWriter

INDONESIAN_SYNONYMS = {
    "menggunakan": "memanfaatkan",
    "memanfaatkan": "menggunakan",
    "menunjukkan": "mengindikasikan",
    "mengindikasikan": "menunjukkan",
    "penelitian": "studi",
    "studi": "penelitian",
    "metode": "pendekatan",
    "pendekatan": "metode",
    "adalah": "merupakan",
    "merupakan": "ialah",
    "berbasis": "berlandaskan",
    "informasi": "data",
    "sistem": "mekanisme",
    "untuk": "guna",
    "sangat": "amat",
    "dapat": "bisa",
    "membuat": "menghasilkan",
    "hasil": "keluaran",
    "analisis": "pengkajian",
    "implementasi": "penerapan",
    "aplikasi": "perangkat lunak",
    "proses": "tahapan",
    "tujuan": "sasaran",
    "maka": "oleh karena itu",
    "sehingga": "akibatnya",
    "namun": "akan tetapi",
    "tetapi": "namun",
    "dengan": "melalui",
    "pada": "di",
    "tentang": "mengenai",
    "secara": "melalui",
    "efisien": "efektif",
    "cepat": "pesat",
    "mudah": "gampang",
    "membantu": "mempermudah",
    "melakukan": "menjalankan",
    "pengembangan": "perancangan",
    "perancangan": "pengembangan",
    "merancang": "mendesain",
    "mendesain": "merancang",
    "menyatakan": "mengungkapkan",
    "mengungkapkan": "menyatakan",
    "menjelaskan": "menerangkan",
    "menerangkan": "menjelaskan",
    "diperlukan": "dibutuhkan",
    "dibutuhkan": "diperlukan",
    "berdasarkan": "berlandaskan",
    "berlandaskan": "berdasarkan",
    "berbagai": "macam-macam",
    "serta": "dan juga",
    "salah satu": "suatu bagian",
    "penting": "krusial",
    "krusial": "penting",
    "teknologi": "sistem teknik",
    "perkembangan": "kemajuan",
    "kemajuan": "perkembangan",
    "meningkatkan": "memaksimalkan",
    "memaksimalkan": "meningkatkan",
    "efektif": "berdaya guna",
    "sebagai": "selaku",
    "karena": "disebabkan oleh",
    "oleh": "lewat"
}

def offline_paraphrase(text):
    words = re.findall(r'\b\w+\b|[^\w\s]', text)
    new_words = []
    for token in words:
        if token.isalnum():
            token_lower = token.lower()
            if token_lower in INDONESIAN_SYNONYMS:
                syn = INDONESIAN_SYNONYMS[token_lower]
                if token.istitle():
                    syn = syn.capitalize()
                elif token.isupper():
                    syn = syn.upper()
                new_words.append(syn)
            else:
                new_words.append(token)
        else:
            new_words.append(token)
            
    result = ""
    for i, token in enumerate(new_words):
        if i > 0 and token.isalnum() and new_words[i-1].isalnum():
            result += " "
        elif i > 0 and token.isalnum() and new_words[i-1] not in ['"', "'", '(', '[', '{']:
            result += " "
        result += token
    return result

app = Flask(__name__)
if os.environ.get('VERCEL'):
    UPLOAD_FOLDER = '/tmp'
else:
    UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Paraphrase System Prompt for Gemini
PARAPHRASE_SYSTEM_PROMPT = """Anda adalah asisten akademik profesional yang berspesialisasi dalam menulis ulang karya ilmiah (skripsi/tesis) untuk menurunkan persentase kemiripan (Turnitin) di tingkat universitas.

Tugas Anda adalah memparafrasekan teks bahasa Indonesia yang diberikan oleh pengguna dengan mengikuti instruksi ketat berikut:

A. Aturan Umum
- Jangan mengubah makna utama, substansi, atau maksud asli penulis sedikit pun.
- Wajib menggunakan bahasa yang sama dengan input teks asli secara mutlak (JANGAN menerjemahkan ke bahasa lain, jika aslinya Bahasa Indonesia maka hasil parafrase wajib Bahasa Indonesia).
- Pertahankan fakta.
- Pertahankan angka.
- Pertahankan statistik.
- Pertahankan hasil penelitian.
- Jangan menambah informasi baru.
- Jangan mengurangi informasi penting.
- Jangan membuat asumsi.
- Jangan membuat opini baru.
- Pertahankan tujuan penulis.

B. Struktur Kalimat
- Variasikan struktur kalimat.
- Hindari pola kalimat yang berulang.
- Pecah kalimat yang terlalu panjang.
- Gabungkan kalimat pendek bila perlu.
- Variasikan posisi subjek.
- Variasikan posisi predikat.
- Variasikan posisi keterangan.
- Hindari pola SPOK yang monoton.
- Gunakan struktur kompleks bila sesuai.
- Gunakan struktur sederhana bila lebih jelas.

C. Diksi
- Gunakan sinonim sesuai konteks.
- Jangan memakai sinonim yang mengubah arti.
- Gunakan istilah akademik.
- Hindari bahasa sehari-hari.
- Hindari kata tidak baku.
- Gunakan kata baku.
- Hindari kata ambigu.
- Gunakan kata yang presisi.
- Hindari pengulangan kata.
- Variasikan pilihan kata.

D. Gaya Akademik
- Gunakan bahasa formal.
- Pertahankan objektivitas.
- Hindari bahasa emosional.
- Hindari hiperbola.
- Hindari kata promosi.
- Hindari bahasa persuasif.
- Gunakan gaya ilmiah.
- Gunakan gaya netral.
- Gunakan istilah konsisten.
- Pertahankan nada akademik.

E. Paragraf
- Pertahankan ide pokok.
- Pertahankan ide pendukung.
- Buat paragraf koheren.
- Gunakan transisi alami.
- Variasikan pembuka paragraf.
- Variasikan penutup paragraf.
- Hindari paragraf terlalu panjang.
- Hindari paragraf satu kalimat bila tidak perlu.
- Hubungkan antar paragraf.
- Pastikan alur logis.

F. Tata Bahasa
- Ikuti PUEBI.
- Perbaiki ejaan.
- Perbaiki kapitalisasi.
- Perbaiki tanda baca.
- Gunakan koma sesuai aturan.
- Gunakan titik dua sesuai aturan.
- Hindari kalimat menggantung.
- Hindari subjek ganda.
- Hindari predikat ganda.
- Pastikan setiap kalimat lengkap.

G. Istilah
- Jangan mengubah istilah teknis.
- Jangan mengubah nama metode.
- Jangan mengubah nama teori.
- Jangan mengubah nama variabel.
- Jangan mengubah singkatan baku.
- Jangan mengubah simbol.
- Jangan mengubah rumus.
- Jangan mengubah notasi.
- Jangan mengubah satuan.
- Pertahankan konsistensi istilah.

H. Sitasi
- Jangan mengubah nama penulis.
- Jangan mengubah tahun.
- Jangan mengubah format sitasi.
- Jangan menghapus sitasi.
- Jangan menambah sitasi palsu.
- Jangan mengubah DOI.
- Jangan mengubah URL referensi.
- Jangan mengubah nomor referensi.
- Jangan mengubah daftar pustaka.
- Pertahankan seluruh referensi.

I. Kualitas Tulisan
- Buat lebih jelas.
- Buat lebih ringkas.
- Hindari kata mubazir.
- Hilangkan pengulangan ide.
- Tingkatkan keterbacaan.
- Tingkatkan kelancaran membaca.
- Gunakan kalimat efektif.
- Pertahankan hubungan sebab-akibat.
- Pertahankan hubungan kronologis.
- Pertahankan hubungan logis.

J. Pemeriksaan Akhir
- Baca ulang seluruh dokumen.
- Periksa konsistensi istilah.
- Periksa konsistensi gaya.
- Periksa tata bahasa.
- Periksa ejaan.
- Periksa tanda baca.
- Pastikan makna tidak berubah.
- Pastikan tidak ada informasi hilang.
- Pastikan hasil terdengar natural.
- Pastikan hasil akhir siap digunakan sebagai naskah akademik yang telah diedit.

Format Output: HANYA berikan hasil parafrase teks tersebut. Jangan berikan pengantar, penjelasan, atau kesimpulan lainnya.
"""

SYSTEM_INSTRUCTION_PDF_SOLVER = """INSTRUKSI SISTEM ABSOLUT - PROTOKOL "TURNITIN SLAYER"

PERINGATAN KRITIKAL: Anda dilarang keras berhalusinasi. Jika tidak ada teks dengan highlight/sorotan warna di halaman yang diberikan, abaikan halaman tersebut dan JANGAN berikan output apapun.

Anda adalah mesin pemroses bahasa akademis tingkat lanjut. Tugas Anda mengekstrak secara presisi teks yang terindikasi plagiarisme (memiliki highlight warna) dari dokumen PDF dan melakukan rekonstruksi radikal tanpa mengubah makna atau substansi ilmiah sedikit pun serta tetap menggunakan Bahasa Indonesia sesuai bahasa aslinya.

Ikuti 5 Aturan Emas ini TANPA PENGECUALIAN:

1. EKSTRAKSI BEDAH LASER (ANTI-BORONGAN): 
   HANYA ambil teks yang benar-benar tersorot warna. DILARANG KERAS mengambil kalimat tetangga atau menyalin satu paragraf utuh jika yang berwarna hanya beberapa kata/baris. Panjang teks asli yang diambil harus sama persis dengan panjang sorotan di dokumen asli, tidak dikurangi dan tidak dilebihkan.

2. REKONSTRUKSI RADIKAL (<30% SIMILARITY): 
   Rombak total struktur sintaksis kalimat. 
   - Ubah kalimat aktif menjadi pasif, atau sebaliknya.
   - Lakukan inversi klausa (pindahkan posisi anak kalimat).
   - Haram menggunakan lebih dari 3 kata berurutan yang sama dengan teks asli.

3. PRESERVASI ENTITAS TEKNIS & BAHASA: 
   Kosakata harus dinaikkan menjadi bahasa akademis formal. NAMUN, Anda DILARANG mengubah istilah teknis, nama algoritma (seperti Haar Cascade, LBPH, dll), bahasa pemrograman, data metrik, angka, dan format sitasi (Nama, Tahun). Wajib memelihara bahasa asli (Bahasa Indonesia) secara mutlak dan dilarang menerjemahkan teks ke bahasa asing.

4. FORMAT MARKDOWN MUTLAK: 
   Wajib menggunakan format keluaran berikut untuk setiap temuan. Jangan tambahkan pembuka, penutup, atau komentar apapun.

### Halaman [Nomor Halaman]
* **Teks Asli:** "[potongan teks asli ber-highlight]"
* **Hasil Parafrase:** "**[kalimat baru yang sudah direkonstruksi]**"

5. EKSEKUSI DIAM:
   Langsung berikan output sesuai format. Jika melanggar, sistem akan menolak respons Anda.
"""

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Slayer Suite</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-color: #f7f7f5;
            --container-bg: #ffffff;
            --text-color: #37352f;
            --border-color: #e9e9e7;
            --hover-bg: #f1f1ef;
            --secondary-text: #787774;
            --accent-color: #000000;
        }

        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
            margin: 0;
            padding: 0;
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: flex-start;
        }

        .container {
            width: 100%;
            max-width: 650px;
            padding: 60px 20px;
            box-sizing: border-box;
        }

        .card {
            background-color: var(--container-bg);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 40px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.02);
        }

        .header {
            margin-bottom: 30px;
            text-align: center;
        }

        h1 {
            font-weight: 700;
            font-size: 2.2rem;
            margin: 0 0 8px 0;
            letter-spacing: -0.03em;
            color: var(--text-color);
        }

        p.subtitle {
            color: var(--secondary-text);
            margin: 0;
            font-size: 1.05rem;
            font-weight: 400;
        }

        .tabs {
            display: flex;
            border-bottom: 1px solid var(--border-color);
            margin-bottom: 30px;
            gap: 16px;
            overflow-x: auto;
        }

        .tab-btn {
            background: none;
            border: none;
            padding: 10px 4px;
            font-size: 0.95rem;
            font-weight: 500;
            color: var(--secondary-text);
            cursor: pointer;
            position: relative;
            transition: color 0.15s ease;
            white-space: nowrap;
        }

        .tab-btn:hover {
            color: var(--text-color);
        }

        .tab-btn.active {
            color: var(--accent-color);
            font-weight: 600;
        }

        .tab-btn.active::after {
            content: '';
            position: absolute;
            bottom: -1px;
            left: 0;
            right: 0;
            height: 2px;
            background-color: var(--accent-color);
        }

        .tab-content {
            display: none;
        }

        .tab-content.active {
            display: block;
        }

        .form-group {
            margin-bottom: 24px;
        }

        label {
            display: block;
            font-weight: 600;
            margin-bottom: 8px;
            font-size: 0.85rem;
            color: var(--text-color);
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .api-input-container {
            display: flex;
            gap: 10px;
            align-items: center;
        }

        input[type="text"] {
            flex-grow: 1;
            padding: 12px;
            background: #ffffff;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            color: var(--text-color);
            font-family: inherit;
            box-sizing: border-box;
            font-size: 0.95rem;
            transition: border-color 0.15s;
        }

        input[type="text"]:focus {
            outline: none;
            border-color: #000000;
        }

        .btn-check-api {
            padding: 12px 20px;
            border: 1px solid var(--border-color);
            background: #ffffff;
            color: var(--text-color);
            font-weight: 600;
            border-radius: 8px;
            cursor: pointer;
            transition: background 0.15s;
            white-space: nowrap;
            box-shadow: 0 1px 2px rgba(15, 15, 15, 0.05);
        }

        .btn-check-api:hover {
            background: var(--hover-bg);
        }

        .api-status-badge {
            margin-top: 8px;
            font-size: 0.85rem;
            font-weight: 600;
            display: none;
        }

        .file-upload-wrapper {
            position: relative;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 24px;
            text-align: center;
            background: #fbfbfa;
            transition: background 0.15s ease, border-color 0.15s ease;
            cursor: pointer;
        }

        .file-upload-wrapper:hover {
            background: var(--hover-bg);
            border-color: #dfdfde;
        }

        .file-upload-wrapper input[type="file"] {
            position: absolute;
            left: 0;
            top: 0;
            opacity: 0;
            width: 100%;
            height: 100%;
            cursor: pointer;
        }

        .upload-icon {
            font-size: 1.8rem;
            margin-bottom: 8px;
            display: block;
        }

        .file-name-label {
            color: var(--secondary-text);
            font-size: 0.9rem;
            font-weight: 500;
            word-break: break-all;
        }

        .btn-submit {
            display: block;
            width: 100%;
            padding: 12px;
            border: 1px solid var(--accent-color);
            border-radius: 8px;
            background: var(--accent-color);
            color: #ffffff;
            font-weight: 600;
            font-size: 1rem;
            cursor: pointer;
            transition: opacity 0.15s ease;
            margin-top: 10px;
        }

        .btn-submit:hover {
            opacity: 0.9;
        }

        .btn-submit.loading {
            background: #e9e9e7;
            border-color: #e9e9e7;
            color: #787774;
            cursor: not-allowed;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }

        .spinner {
            width: 16px;
            height: 16px;
            border: 2px solid #787774;
            border-top-color: transparent;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            display: inline-block;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .btn-download {
            display: inline-block;
            padding: 12px 24px;
            border: 1px solid var(--accent-color);
            border-radius: 8px;
            background: var(--accent-color);
            color: #ffffff;
            font-weight: 600;
            font-size: 0.95rem;
            text-decoration: none;
            transition: opacity 0.15s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .btn-download:hover {
            opacity: 0.9;
        }

        .alert-box {
            background: #f1f8f5;
            border: 1px solid #d4ebdf;
            color: #0f5132;
            border-radius: 8px;
            padding: 14px;
            margin-bottom: 25px;
            font-size: 0.9rem;
            font-weight: 500;
        }

        .error-box {
            background: #fdf2f2;
            border: 1px solid #fde8e8;
            color: #9b1c1c;
            border-radius: 8px;
            padding: 14px;
            margin-bottom: 25px;
            font-size: 0.9rem;
            font-weight: 500;
            word-wrap: break-word;
        }

        .rules-container {
            max-height: 150px;
            overflow-y: auto;
            border: 1px solid var(--border-color);
            background: #fbfbfa;
            padding: 10px 15px;
            border-radius: 8px;
            font-size: 0.8rem;
            color: var(--secondary-text);
        }
        
        .rules-container ul {
            margin: 0;
            padding-left: 20px;
        }

        footer {
            margin-top: 40px;
            font-size: 0.8rem;
            color: var(--secondary-text);
            text-align: center;
        }

        .main-layout {
            display: grid;
            grid-template-columns: 1fr 340px;
            gap: 25px;
            max-width: 1200px;
            margin: 40px auto;
            padding: 0 20px;
            box-sizing: border-box;
        }

        @media (max-width: 950px) {
            .main-layout {
                grid-template-columns: 1fr;
            }
        }

        .sidebar-card {
            background: #ffffff;
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 1px 3px rgba(15, 15, 15, 0.05);
            max-height: 80vh;
            overflow-y: auto;
            position: sticky;
            top: 40px;
        }

        .sidebar-header h2 {
            font-size: 1.05rem;
            margin: 0 0 6px 0;
            font-weight: 700;
        }

        .sidebar-header p {
            margin: 0 0 20px 0;
            font-size: 0.8rem;
            color: var(--secondary-text);
            line-height: 1.4;
        }

        .sidebar-section h3 {
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--secondary-text);
            margin: 0 0 10px 0;
            font-weight: 700;
        }

        .file-list {
            display: flex;
            flex-direction: column;
            gap: 8px;
        }

        .file-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #fbfbfa;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            padding: 8px 12px;
            font-size: 0.8rem;
        }

        .file-name {
            font-weight: 500;
            color: var(--text-color);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            max-width: 190px;
        }

        .file-actions {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .empty-files {
            font-size: 0.8rem;
            color: var(--secondary-text);
            font-style: italic;
            text-align: center;
            padding: 10px 0;
        }
    </style>
</head>
<body>
    <div class="main-layout">
        <div class="container" style="margin: 0; padding: 0; width: 100%; max-width: 100%;">
            <div class="card">
            <div class="header">
                <h1>Slayer Suite</h1>
                <p class="subtitle">Turnitin Bypass & PDF Productivity Tools</p>
            </div>

            <div class="tabs">
                <button class="tab-btn active" onclick="switchTab('manual')">Turnitin Manual</button>
                <button class="tab-btn" onclick="switchTab('ai')">Turnitin AI Pro</button>
                <button class="tab-btn" onclick="switchTab('compress')">PDF Compressor</button>
                <button class="tab-btn" onclick="switchTab('pdf2word')">PDF to Word</button>
                <button class="tab-btn" onclick="switchTab('word2pdf')">Word to PDF</button>
                <button class="tab-btn" onclick="switchTab('pdf_solver')">PDF Solver (AI)</button>
                <button class="tab-btn" onclick="switchTab('compare_docs')">Compare Docs</button>
            </div>

            {% if success_msg %}
                <div class="alert-box">{{ success_msg }}</div>
            {% endif %}

            {% if error_msg %}
                <div class="error-box">{{ error_msg }}</div>
            {% endif %}

            {% if result_file %}
                <div style="margin-bottom: 30px; text-align: center; background: #fbfbfa; padding: 20px; border: 1px solid var(--border-color); border-radius: 8px;">
                    <p style="margin: 0 0 12px 0; font-size: 0.9rem; font-weight: 500; color: var(--text-color);">Proses Selesai! Hasil Anda siap diunduh.</p>
                    <a href="/download/{{ result_file }}" class="btn-download">📥 Download Result Document</a>
                </div>
            {% endif %}

            <!-- TAB 1: TURNITIN MANUAL -->
            <div id="manual" class="tab-content active">
                <form action="/process_manual" method="POST" enctype="multipart/form-data">
                    {% if has_active_replacements %}
                    <div style="background: #fdfdfd; border: 1px solid #e9e9e7; padding: 16px; border-radius: 8px; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 1px 3px rgba(0,0,0,0.02);">
                        <div>
                            <strong style="font-size: 0.9rem; color: #1a1a1a; display: flex; align-items: center; gap: 6px;">⚡ Sesi Revisi Aktif Ditemukan</strong>
                            <p style="margin: 4px 0 0 0; font-size: 0.8rem; color: var(--secondary-text);">Ada {{ active_replacements_count }} kata/kalimat hasil parafrase PDF Solver siap diaplikasikan.</p>
                        </div>
                        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 0.85rem; font-weight: 600; color: #1a1a1a; padding: 6px 12px; background: #f1f1ef; border-radius: 6px; border: 1px solid var(--border-color);">
                            <input type="checkbox" name="use_active_session" id="use_active_session" onchange="toggleManualRefUpload()" style="margin: 0;"> Gunakan Sesi
                        </label>
                    </div>
                    {% endif %}

                    <div class="form-group">
                        <label>Original Skripsi Document (.docx)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">📄</span>
                            <div class="file-name-label" id="manual-orig-label">Choose a file or drag it here</div>
                            <input type="file" name="original_doc" accept=".docx" required onchange="updateLabel(this, 'manual-orig-label')">
                        </div>
                    </div>

                    <div class="form-group" id="manual_ref_group">
                        <label>Paraphrase Reference Document (.docx)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">✍️</span>
                            <div class="file-name-label" id="manual-ref-label">Choose a file or drag it here</div>
                            <input type="file" name="reference_doc" id="manual_ref_input" accept=".docx" required onchange="updateLabel(this, 'manual-ref-label')">
                        </div>
                    </div>

                    <button type="submit" class="btn-submit">Process & Replace</button>
                </form>
            </div>

            <!-- TAB 2: TURNITIN AI PRO -->
            <div id="ai" class="tab-content">
                <form action="/process_ai" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Paraphrase Engine</label>
                        <select name="engine" id="engine_select" onchange="toggleAPIKeyField()" style="width: 100%; padding: 10px; border: 1px solid var(--border-color); border-radius: 6px; font-family: inherit; font-size: 0.9rem; background: #fbfbfa; color: var(--text-color); margin-bottom: 15px;">
                            <option value="gemini">Gemini AI (Requires Google Key)</option>
                            <option value="openrouter">OpenRouter AI (Free & Credits Key compatible)</option>
                            <option value="local">Local Paraphraser (100% Free & Offline - No Key Required)</option>
                        </select>
                    </div>

                    <div class="form-group" id="api_key_group">
                        <label>Gemini / OpenRouter API Key</label>
                        <div class="api-input-container">
                            <input type="text" id="api_key_input" name="api_key" placeholder="Enter your API Key here" required value="{{ api_key_val }}">
                            <button type="button" class="btn-check-api" onclick="checkAPIKey()">Check API</button>
                        </div>
                        <div id="api_status" class="api-status-badge"></div>
                    </div>

                    <div class="form-group">
                        <label>Skripsi Document to Paraphrase (.docx)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">📄</span>
                            <div class="file-name-label" id="ai-orig-label">Choose a file or drag it here</div>
                            <input type="file" name="original_doc" accept=".docx" required onchange="updateLabel(this, 'ai-orig-label')">
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Paraphrase Rules (Active & Updated)</label>
                        <div class="rules-container">
                            <ul>
                                <li><strong>Aturan Umum:</strong> Jangan mengubah makna utama, pertahankan fakta/angka/statistik/hasil penelitian.</li>
                                <li><strong>Struktur Kalimat:</strong> Variasikan struktur, hindari pola kalimat berulang, variasikan posisi S-P-O-K.</li>
                                <li><strong>Diksi & Gaya Akademik:</strong> Gunakan sinonim kontekstual yang presisi, bahasa formal objektif netral.</li>
                                <li><strong>Paragraf & Istilah:</strong> Transisi alami. Jangan ubah istilah teknis, nama teori/metode/variabel.</li>
                                <li><strong>Sitasi & Kualitas:</strong> Pertahankan nama penulis/tahun/format sitasi, ejaan sesuai PUEBI.</li>
                            </ul>
                        </div>
                    </div>

                    <button type="submit" class="btn-submit">Slay Turnitin with AI</button>
                </form>
            </div>

            <!-- TAB 3: PDF COMPRESSOR -->
            <div id="compress" class="tab-content">
                <form action="/process_compress" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Select PDF File to Compress</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">📦</span>
                            <div class="file-name-label" id="compress-pdf-label">Choose a file or drag it here</div>
                            <input type="file" name="pdf_file" accept=".pdf" required onchange="updateLabel(this, 'compress-pdf-label')">
                        </div>
                    </div>

                    <button type="submit" class="btn-submit">Compress PDF</button>
                </form>
            </div>

            <!-- TAB 4: PDF TO WORD -->
            <div id="pdf2word" class="tab-content">
                <form action="/process_convert" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Select PDF File to Convert</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">📕</span>
                            <div class="file-name-label" id="convert-pdf-label">Choose a file or drag it here</div>
                            <input type="file" name="pdf_file" accept=".pdf" required onchange="updateLabel(this, 'convert-pdf-label')">
                        </div>
                    </div>

                    <button type="submit" class="btn-submit">Convert to Word</button>
                </form>
            </div>

            <!-- TAB 5: WORD TO PDF -->
            <div id="word2pdf" class="tab-content">
                <form action="/process_word2pdf" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Select Word File to Convert (.docx)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">📄</span>
                            <div class="file-name-label" id="word2pdf-doc-label">Choose a file or drag it here</div>
                            <input type="file" name="doc_file" accept=".docx" required onchange="updateLabel(this, 'word2pdf-doc-label')">
                        </div>
                    </div>

                    <button type="submit" class="btn-submit">Convert to PDF</button>
                </form>
            </div>

            <!-- TAB 6: TURNITIN PDF SOLVER -->
            <div id="pdf_solver" class="tab-content">
                <form action="/process_pdf_solver" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Paraphrase Engine</label>
                        <select name="engine" id="engine_select_solver" onchange="toggleAPIKeyFieldSolver()" style="width: 100%; padding: 10px; border: 1px solid var(--border-color); border-radius: 6px; font-family: inherit; font-size: 0.9rem; background: #fbfbfa; color: var(--text-color); margin-bottom: 15px;">
                            <option value="gemini">Gemini AI (Requires Google Key)</option>
                            <option value="openrouter">OpenRouter AI (Free & Credits Key compatible)</option>
                            <option value="local">Local Paraphraser (100% Free & Offline - No Key Required)</option>
                        </select>
                    </div>

                    <div class="form-group" id="api_key_group_solver">
                        <label>Gemini / OpenRouter API Key</label>
                        <div class="api-input-container">
                            <input type="text" id="api_key_input_solver" name="api_key" placeholder="Enter your API Key here" required value="{{ api_key_val }}">
                            <button type="button" class="btn-check-api" onclick="checkAPIKeySolver()">Check API</button>
                        </div>
                        <div id="api_status_solver" class="api-status-badge"></div>
                    </div>

                    <div class="form-group">
                        <label>Turnitin PDF Result with Highlights (.pdf)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">📕</span>
                            <div class="file-name-label" id="pdf-solver-label">Choose a file or drag it here</div>
                            <input type="file" name="pdf_file" accept=".pdf" required onchange="updateLabel(this, 'pdf-solver-label')">
                        </div>
                    </div>

                    <button type="submit" class="btn-submit">Analyze & Slay PDF</button>
                </form>
            </div>

            <!-- TAB 7: COMPARE DOCS -->
            <div id="compare_docs" class="tab-content">
                <form action="/process_compare_docs" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Original Skripsi Document (.docx)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">📄</span>
                            <div class="file-name-label" id="compare-orig-label">Choose original file or drag it here</div>
                            <input type="file" name="original_doc" accept=".docx" required onchange="updateLabel(this, 'compare-orig-label')">
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Paraphrased Skripsi Document (.docx)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">✍️</span>
                            <div class="file-name-label" id="compare-para-label">Choose paraphrased file or drag it here</div>
                            <input type="file" name="paraphrased_doc" accept=".docx" required onchange="updateLabel(this, 'compare-para-label')">
                        </div>
                    </div>

                    <button type="submit" class="btn-submit">Compare Documents</button>
                </form>
            </div>

            <footer>Powered by Antigravity AI</footer>
        </div>
    </div>

    <!-- Side Storage Panel -->
    <div class="sidebar-card">
        <div class="sidebar-header">
            <h2>📁 Server Storage</h2>
            <p>Klik tombol panah (➡️) untuk mengisi berkas di server ke kolom input tab aktif tanpa upload ulang.</p>
        </div>
        
        <div class="sidebar-section">
            <h3>Word Documents (.docx)</h3>
            <div class="file-list">
                {% if docx_files %}
                    {% for file in docx_files %}
                    <div class="file-item">
                        <span class="file-name" title="{{ file }}">{{ file }}</span>
                        <div class="file-actions">
                            <a href="/download/{{ file }}" title="Download" style="text-decoration:none;">📥</a>
                            <button type="button" onclick="showFileMenu(event, '{{ file }}', true)" title="Masukkan berkas" style="background:none; border:none; cursor:pointer; font-size:1.1rem; padding:0; line-height:1; display:inline-flex; align-items:center;">➡️</button>
                        </div>
                    </div>
                    {% endfor %}
                {% else %}
                    <div class="empty-files">Belum ada berkas DOCX di server.</div>
                {% endif %}
            </div>
        </div>

        <div class="sidebar-section" style="margin-top: 25px;">
            <h3>PDF Documents (.pdf)</h3>
            <div class="file-list">
                {% if pdf_files %}
                    {% for file in pdf_files %}
                    <div class="file-item">
                        <span class="file-name" title="{{ file }}">{{ file }}</span>
                        <div class="file-actions">
                            <a href="/download/{{ file }}" title="Download" style="text-decoration:none;">📥</a>
                            <button type="button" onclick="showFileMenu(event, '{{ file }}', false)" title="Masukkan berkas" style="background:none; border:none; cursor:pointer; font-size:1.1rem; padding:0; line-height:1; display:inline-flex; align-items:center;">➡️</button>
                        </div>
                    </div>
                    {% endfor %}
                {% else %}
                    <div class="empty-files">Belum ada berkas PDF di server.</div>
                {% endif %}
            </div>
        </div>
    </div>
</div>

    <script>
        function updateLabel(input, labelId) {
            const label = document.getElementById(labelId);
            if (input.files && input.files.length > 0) {
                label.innerText = input.files[0].name;
                label.style.color = 'var(--text-color)';
            } else {
                label.innerText = 'Choose a file or drag it here';
                label.style.color = 'var(--secondary-text)';
            }
        }

        function switchTab(tabId) {
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
            
            const activeBtn = Array.from(document.querySelectorAll('.tab-btn')).find(btn => btn.getAttribute('onclick').includes(tabId));
            if (activeBtn) activeBtn.classList.add('active');
            
            const activeContent = document.getElementById(tabId);
            if (activeContent) activeContent.classList.add('active');

            // Save tab state in URL hash
            window.location.hash = tabId;
        }

        function toggleAPIKeyField() {
            const engine = document.getElementById('engine_select').value;
            const apiKeyContainer = document.getElementById('api_key_group');
            const apiKeyInput = document.getElementById('api_key_input');
            if (engine === 'local') {
                apiKeyContainer.style.display = 'none';
                apiKeyInput.removeAttribute('required');
            } else {
                apiKeyContainer.style.display = 'block';
                apiKeyInput.setAttribute('required', 'required');
            }
        }

        function toggleAPIKeyFieldSolver() {
            const engine = document.getElementById('engine_select_solver').value;
            const apiKeyContainer = document.getElementById('api_key_group_solver');
            const apiKeyInput = document.getElementById('api_key_input_solver');
            if (engine === 'local') {
                apiKeyContainer.style.display = 'none';
                apiKeyInput.removeAttribute('required');
            } else {
                apiKeyContainer.style.display = 'block';
                apiKeyInput.setAttribute('required', 'required');
            }
        }

        function toggleManualRefUpload() {
            const checkbox = document.getElementById('use_active_session');
            const refGroup = document.getElementById('manual_ref_group');
            const refInput = document.getElementById('manual_ref_input');
            if (checkbox && checkbox.checked) {
                if (refGroup) refGroup.style.display = 'none';
                if (refInput) refInput.removeAttribute('required');
            } else {
                if (refGroup) refGroup.style.display = 'block';
                if (refInput) refInput.setAttribute('required', 'required');
            }
        }

        // Initialize active tab based on Hash or parameter
        window.addEventListener('DOMContentLoaded', () => {
            const urlParams = new URLSearchParams(window.location.search);
            const activeTabParam = urlParams.get('active_tab');
            const hash = window.location.hash.substring(1);
            
            if (activeTabParam) {
                switchTab(activeTabParam);
            } else if (hash) {
                switchTab(hash);
            }
            toggleAPIKeyField();
            toggleAPIKeyFieldSolver();
            toggleManualRefUpload();
        });

        function checkAPIKey() {
            const apiKey = document.getElementById('api_key_input').value.trim();
            const statusDiv = document.getElementById('api_status');
            
            if (!apiKey) {
                statusDiv.style.display = 'block';
                statusDiv.style.color = '#9b1c1c';
                statusDiv.innerText = '⚠️ Silakan masukkan API Key terlebih dahulu!';
                return;
            }

            statusDiv.style.display = 'block';
            statusDiv.style.color = 'var(--secondary-text)';
            statusDiv.innerText = '⌛ Memeriksa status API Key...';

            fetch('/check_api', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ api_key: apiKey })
            })
            .then(response => response.json())
            .then(data => {
                if (data.status === 'success') {
                    statusDiv.style.color = '#0f5132';
                    statusDiv.innerText = '🟢 API Key Aktif & Valid! Siap digunakan.';
                } else {
                    statusDiv.style.color = '#9b1c1c';
                    statusDiv.innerText = '🔴 API Key Tidak Valid / Error: ' + data.message;
                }
            })
            .catch(error => {
                statusDiv.style.color = '#9b1c1c';
                statusDiv.innerText = '🔴 Gagal terhubung ke server untuk mengecek API Key.';
            });
        }

        function checkAPIKeySolver() {
            const apiKey = document.getElementById('api_key_input_solver').value.strip ? document.getElementById('api_key_input_solver').value.strip() : document.getElementById('api_key_input_solver').value;
            const statusDiv = document.getElementById('api_status_solver');
            
            if (!apiKey) {
                statusDiv.style.display = 'block';
                statusDiv.style.color = '#9b1c1c';
                statusDiv.innerText = '⚠️ Silakan masukkan API Key terlebih dahulu!';
                return;
            }

            statusDiv.style.display = 'block';
            statusDiv.style.color = 'var(--secondary-text)';
            statusDiv.innerText = '⌛ Memeriksa status API Key...';

            fetch('/check_api', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ api_key: apiKey })
            })
            .then(response => response.json())
            .then(data => {
                if (data.status === 'success') {
                    statusDiv.style.color = '#0f5132';
                    statusDiv.innerText = '🟢 API Key Aktif & Valid! Siap digunakan.';
                } else {
                    statusDiv.style.color = '#9b1c1c';
                    statusDiv.innerText = '🔴 API Key Tidak Valid / Error: ' + data.message;
                }
            })
            .catch(error => {
                statusDiv.style.color = '#9b1c1c';
                statusDiv.innerText = '🔴 Gagal terhubung ke server untuk mengecek API Key.';
            });
        }

        function showFileMenu(e, filename, isDocx) {
            e.stopPropagation();
            const existing = document.querySelector('.file-action-menu');
            if (existing) existing.remove();

            const menu = document.createElement('div');
            menu.className = 'file-action-menu';
            menu.style.position = 'absolute';
            menu.style.background = '#ffffff';
            menu.style.border = '1px solid var(--border-color)';
            menu.style.borderRadius = '8px';
            menu.style.boxShadow = '0 4px 12px rgba(0,0,0,0.1)';
            menu.style.zIndex = '1000';
            menu.style.padding = '8px 0';
            menu.style.minWidth = '220px';
            menu.style.top = (e.pageY + 10) + 'px';
            menu.style.left = (e.pageX - 180) + 'px';

            let items = [];
            if (isDocx) {
                items = [
                    { text: '📄 Set as Original (Manual)', action: () => selectServerFile('manual', 'original_doc', filename) },
                    { text: '✍️ Set as Reference (Manual)', action: () => selectServerFile('manual', 'reference_doc', filename) },
                    { text: '🤖 Set as Original (AI Pro)', action: () => selectServerFile('ai', 'original_doc', filename) },
                    { text: '📄 Set as Original (Compare)', action: () => selectServerFile('compare_docs', 'original_doc', filename) },
                    { text: '✍️ Set as Paraphrased (Compare)', action: () => selectServerFile('compare_docs', 'paraphrased_doc', filename) }
                ];
            } else {
                items = [
                    { text: '📕 Set as PDF Solver Input', action: () => selectServerFile('pdf_solver', 'pdf_file', filename) }
                ];
            }

            items.forEach(item => {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.style.width = '100%';
                btn.style.padding = '8px 16px';
                btn.style.textAlign = 'left';
                btn.style.background = 'none';
                btn.style.border = 'none';
                btn.style.cursor = 'pointer';
                btn.style.fontSize = '0.8rem';
                btn.style.fontFamily = 'inherit';
                btn.innerText = item.text;
                btn.onmouseenter = () => btn.style.background = '#f5f5f4';
                btn.onmouseleave = () => btn.style.background = 'none';
                btn.onclick = () => {
                    item.action();
                    menu.remove();
                };
                menu.appendChild(btn);
            });

            document.body.appendChild(menu);
            
            const closeHandler = () => {
                menu.remove();
                document.removeEventListener('click', closeHandler);
            };
            setTimeout(() => {
                document.addEventListener('click', closeHandler);
            }, 10);
        }

        function selectServerFile(tabId, fieldName, filename) {
            switchTab(tabId);
            
            const form = document.querySelector(`#${tabId} form`);
            if (!form) return;
            
            let hiddenInput = form.querySelector(`input[name="server_${fieldName}"]`);
            if (!hiddenInput) {
                hiddenInput = document.createElement('input');
                hiddenInput.type = 'hidden';
                hiddenInput.name = `server_${fieldName}`;
                form.appendChild(hiddenInput);
            }
            hiddenInput.value = filename;
            
            let labelId = '';
            if (tabId === 'manual' && fieldName === 'original_doc') labelId = 'manual-orig-label';
            if (tabId === 'manual' && fieldName === 'reference_doc') labelId = 'manual-ref-label';
            if (tabId === 'ai' && fieldName === 'original_doc') labelId = 'ai-orig-label';
            if (tabId === 'compare_docs' && fieldName === 'original_doc') labelId = 'compare-orig-label';
            if (tabId === 'compare_docs' && fieldName === 'paraphrased_doc') labelId = 'compare-para-label';
            if (tabId === 'pdf_solver' && fieldName === 'pdf_file') labelId = 'pdf-solver-label';
            
            if (labelId) {
                const label = document.getElementById(labelId);
                if (label) {
                    label.innerHTML = `<span style="color: #0f5132; font-weight: 600;">📁 Server: ${filename}</span>`;
                }
            }
            
            const fileInput = form.querySelector(`input[name="${fieldName}"]`);
            if (fileInput) {
                fileInput.removeAttribute('required');
            }
        }

        // Show loading state on form submit
        window.addEventListener('DOMContentLoaded', () => {
            document.querySelectorAll('form').forEach(form => {
                form.addEventListener('submit', function(e) {
                    const submitBtn = form.querySelector('.btn-submit');
                    if (submitBtn) {
                        submitBtn.style.pointerEvents = 'none';
                        submitBtn.style.opacity = '0.7';
                        submitBtn.innerHTML = '<div class="spinner"></div> Mohon tunggu, sedang memproses...';
                    }
                });
            });
        });
    </script>
</body>
</html>
"""

def parse_revisi(path):
    doc = Document(path)
    replacements = []
    current_original = None
    current_paraphrase = None
    
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        
        orig_match = re.search(r'\*\*Teks Asli:\*\*\s*(.*)$', txt)
        para_match = re.search(r'\*\*Hasil Parafrase:\*\*\s*(.*)$', txt)
        
        if orig_match:
            val = orig_match.group(1).strip()
            if val.startswith('"') and val.endswith('"'):
                val = val[1:-1]
            elif val.startswith('“') and val.endswith('”'):
                val = val[1:-1]
            elif val.startswith("'") and val.endswith("'"):
                val = val[1:-1]
            current_original = val.strip()
        elif para_match:
            val = para_match.group(1).strip()
            if val.startswith('**'):
                val = val[2:]
            if val.endswith('**'):
                val = val[:-2]
            val = val.strip()
            if val.startswith('"') and val.endswith('"'):
                val = val[1:-1]
            elif val.startswith('“') and val.endswith('”'):
                val = val[1:-1]
            elif val.startswith("'") and val.endswith("'"):
                val = val[1:-1]
            current_paraphrase = val.strip()
            
            if current_original:
                current_paraphrase = current_paraphrase.replace("**", "")
                replacements.append((current_original, current_paraphrase))
                current_original = None
                current_paraphrase = None
                
    return replacements

def do_smart_replacements(original_doc_path, replacements, output_path):
    doc = Document(original_doc_path)
    
    def clean(text):
        return re.sub(r'\s+', ' ', text).strip()
        
    sorted_reps = sorted(replacements, key=lambda x: len(x[0]), reverse=True)
    replaced_count = 0
    
    for p in doc.paragraphs:
        p_clean = clean(p.text)
        for orig, para in sorted_reps:
            orig_clean = clean(orig)
            if orig_clean in p_clean:
                if len(p_clean) - len(orig_clean) < 10:
                    if len(p.runs) > 0:
                        p.runs[0].text = para
                        for r in p.runs[1:]:
                            r.text = ""
                    else:
                        p.text = para
                    replaced_count += 1
                    p_clean = clean(para)
                else:
                    if orig in p.text:
                        p.text = p.text.replace(orig, para)
                        replaced_count += 1
                        p_clean = clean(p.text)
                    elif orig_clean in clean(p.text):
                        words = [re.escape(w) for w in orig_clean.split()]
                        pattern = r'\s+'.join(words)
                        new_text, count = re.subn(pattern, para, p.text)
                        if count > 0:
                            p.text = new_text
                            replaced_count += count
                            p_clean = clean(p.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p_clean = clean(p.text)
                    for orig, para in sorted_reps:
                        orig_clean = clean(orig)
                        if orig_clean in p_clean:
                            if len(p_clean) - len(orig_clean) < 10:
                                if len(p.runs) > 0:
                                    p.runs[0].text = para
                                    for r in p.runs[1:]:
                                        r.text = ""
                                else:
                                    p.text = para
                                replaced_count += 1
                                p_clean = clean(para)
                            else:
                                if orig in p.text:
                                    p.text = p.text.replace(orig, para)
                                    replaced_count += 1
                                    p_clean = clean(p.text)
                                elif orig_clean in clean(p.text):
                                    words = [re.escape(w) for w in orig_clean.split()]
                                    pattern = r'\s+'.join(words)
                                    new_text, count = re.subn(pattern, para, p.text)
                                    if count > 0:
                                        p.text = new_text
                                        replaced_count += count
                                        p_clean = clean(p.text)
                                        
    doc.save(output_path)
    return replaced_count

@app.route('/')
def index():
    success_msg = request.args.get('success_msg')
    error_msg = request.args.get('error_msg')
    result_file = request.args.get('result_file')
    api_key_val = request.args.get('api_key_val', '')
    
    import json
    active_reps_path = os.path.join(app.config['UPLOAD_FOLDER'], 'active_replacements.json')
    has_active_replacements = False
    active_replacements_count = 0
    if os.path.exists(active_reps_path):
        try:
            with open(active_reps_path, 'r') as f:
                reps = json.load(f)
                if reps:
                    has_active_replacements = True
                    active_replacements_count = len(reps)
        except:
            pass
            
    # List files in upload folder for the side select storage panel
    docx_files = []
    pdf_files = []
    try:
        upload_dir = app.config['UPLOAD_FOLDER']
        if os.path.exists(upload_dir):
            for file in os.listdir(upload_dir):
                if file.startswith('.') or file == 'active_replacements.json':
                    continue
                path = os.path.join(upload_dir, file)
                if os.path.isfile(path):
                    if file.lower().endswith('.docx'):
                        docx_files.append(file)
                    elif file.lower().endswith('.pdf'):
                        pdf_files.append(file)
    except:
        pass

    return render_template_string(HTML_TEMPLATE, 
                                  success_msg=success_msg, 
                                  error_msg=error_msg, 
                                  result_file=result_file, 
                                  api_key_val=api_key_val,
                                  has_active_replacements=has_active_replacements,
                                  active_replacements_count=active_replacements_count,
                                  docx_files=docx_files,
                                  pdf_files=pdf_files)

@app.route('/check_api', methods=['POST'])
def check_api():
    data = request.get_json()
    api_key = data.get('api_key')
    if not api_key:
        return jsonify({'status': 'error', 'message': 'API Key kosong'})
        
    try:
        if api_key.startswith('sk-or-'):
            import requests
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": "openrouter/free",
                "messages": [{"role": "user", "content": "Tes"}],
                "max_tokens": 10
            }
            res = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)
            if res.status_code == 200:
                return jsonify({'status': 'success'})
            else:
                return jsonify({'status': 'error', 'message': f'OpenRouter Error: {res.text}'})
        else:
            genai.configure(api_key=api_key)
            models_to_try = [
                'gemini-2.0-flash-latest', 'gemini-2.0-flash', 
                'gemini-2.5-flash', 'gemini-2.5-flash-lite', 
                'gemini-1.5-flash', 'gemini-1.5-flash-latest', 
                'gemini-1.5-pro', 'gemini-2.5-pro'
            ]
            response = None
            last_err = None
            
            for model_name in models_to_try:
                try:
                    model = genai.GenerativeModel(model_name)
                    response = model.generate_content("Tes koneksi. Balas dengan kata OK saja.")
                    if response and response.text:
                        break
                except Exception as err:
                    last_err = err
                    
            if response and response.text:
                return jsonify({'status': 'success'})
            else:
                return jsonify({'status': 'error', 'message': f'Respons kosong atau error dari API. Terakhir error: {last_err}'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

def resolve_file(form_server_param, file_key, allowed_ext):
    server_file = request.form.get(form_server_param)
    if server_file:
        filename = server_file
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if not os.path.exists(path):
            raise ValueError(f"File server '{filename}' tidak ditemukan.")
        return filename, path, False
    else:
        if file_key not in request.files:
            raise ValueError(f"Silakan pilih file untuk '{file_key}'.")
        f = request.files[file_key]
        if f.filename == '':
            raise ValueError("Nama file tidak valid.")
        if not f.filename.lower().endswith(allowed_ext):
            raise ValueError(f"Format file harus berupa {allowed_ext}!")
        path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
        f.save(path)
        return f.filename, path, True

@app.route('/process_manual', methods=['POST'])
def process_manual():
    use_active_session = request.form.get('use_active_session') == 'on'
    
    try:
        orig_name, orig_path, is_temp_orig = resolve_file('server_original_doc', 'original_doc', '.docx')
    except Exception as e:
        return redirect(url_for('index', error_msg=str(e), active_tab="manual"))
        
    reps = []
    active_reps_path = os.path.join(app.config['UPLOAD_FOLDER'], 'active_replacements.json')
    
    if use_active_session:
        import json
        if os.path.exists(active_reps_path):
            try:
                with open(active_reps_path, 'r') as f:
                    reps = json.load(f)
            except Exception as e:
                return redirect(url_for('index', error_msg=f"Error reading session replacements: {str(e)}", active_tab="manual"))
        if not reps:
            return redirect(url_for('index', error_msg="No active session replacements found.", active_tab="manual"))
    else:
        try:
            ref_name, ref_path, is_temp_ref = resolve_file('server_reference_doc', 'reference_doc', '.docx')
            reps = parse_revisi(ref_path)
            if is_temp_ref:
                try: os.remove(ref_path)
                except: pass
        except Exception as e:
            if is_temp_orig:
                try: os.remove(orig_path)
                except: pass
            return redirect(url_for('index', error_msg=str(e), active_tab="manual"))
            
    if reps:
        try:
            out_filename = "PARAFRASED_" + orig_name
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            
            replaced_count = do_smart_replacements(orig_path, reps, out_path)
            
            if is_temp_orig:
                try: os.remove(orig_path)
                except: pass
                
            if use_active_session:
                try: os.remove(active_reps_path)
                except: pass
                
            return redirect(url_for('index', success_msg=f"Success! Replaced {replaced_count} matches.", result_file=out_filename, active_tab="manual"))
        except Exception as e:
            return redirect(url_for('index', error_msg=f"Error occurred during replacement: {str(e)}", active_tab="manual"))
            
    return redirect(url_for('index', error_msg="Unknown error.", active_tab="manual"))

@app.route('/process_ai', methods=['POST'])
def process_ai():
    engine = request.form.get('engine', 'gemini')
    api_key = request.form.get('api_key')
    
    if engine != 'local' and not api_key:
        return redirect(url_for('index', error_msg="API Key is required for AI modes.", active_tab="ai"))
        
    try:
        orig_name, orig_path, is_temp_orig = resolve_file('server_original_doc', 'original_doc', '.docx')
    except Exception as e:
        return redirect(url_for('index', error_msg=str(e), api_key_val=api_key, active_tab="ai"))
    
    try:
        try:
            doc = Document(orig_path)
        except Exception as docx_err:
            raise Exception(f"File dokumen tidak bisa dibuka atau rusak: {docx_err}")

        replaced_count = 0
        
        for p in doc.paragraphs:
            original_text = p.text.strip()
            if len(original_text) > 20 and not original_text.startswith("BAB ") and not original_text.startswith("DAFTAR PUSTAKA") and not original_text.isupper():
                try:
                    para_text = None
                    last_err = None
                    
                    if engine == 'local':
                        para_text = offline_paraphrase(original_text)
                    elif engine == 'openrouter':
                        import requests
                        headers = {
                            "Authorization": f"Bearer {api_key}",
                            "Content-Type": "application/json"
                        }
                        payload = {
                            "model": "openrouter/free",
                            "messages": [
                                {"role": "system", "content": PARAPHRASE_SYSTEM_PROMPT},
                                {"role": "user", "content": original_text}
                            ],
                            "max_tokens": 1000
                        }
                        res = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)
                        if res.status_code == 200:
                            para_text = res.json()["choices"][0]["message"]["content"].strip()
                        else:
                            last_err = res.text
                    else:
                        genai.configure(api_key=api_key)
                        models_to_try = [
                            'gemini-2.0-flash-latest', 'gemini-2.0-flash', 
                            'gemini-2.5-flash', 'gemini-2.5-flash-lite', 
                            'gemini-1.5-flash', 'gemini-1.5-flash-latest', 
                            'gemini-1.5-pro', 'gemini-2.5-pro'
                        ]
                        for model_name in models_to_try:
                            try:
                                model = genai.GenerativeModel(model_name, system_instruction=PARAPHRASE_SYSTEM_PROMPT)
                                response = model.generate_content(original_text)
                                if response and response.text:
                                    para_text = response.text.strip()
                                    break
                            except Exception as m_err:
                                last_err = m_err
                            
                    if not para_text:
                        print(f"Skipping paragraph due to API failure. Last error: {last_err}")
                        continue
                    
                    if len(p.runs) > 0:
                        p.runs[0].text = para_text
                        for r in p.runs[1:]:
                            r.text = ""
                    else:
                        p.text = para_text
                    replaced_count += 1
                except Exception as api_err:
                    print(f"API Error at paragraph: {original_text[:50]}... Error: {api_err}")
                    
        out_filename = "AI_PARAFRASED_" + orig_name
        out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
        doc.save(out_path)
        
        if is_temp_orig:
            try:
                os.remove(orig_path)
            except:
                pass
            
        return redirect(url_for('index', success_msg=f"Successfully paraphrased {replaced_count} paragraphs!", result_file=out_filename, api_key_val=api_key, active_tab="ai"))
        
    except Exception as e:
        try:
            if os.path.exists(orig_path):
                os.remove(orig_path)
        except:
            pass
        return redirect(url_for('index', error_msg=f"Error occurred: {str(e)}", api_key_val=api_key, active_tab="ai"))

@app.route('/process_compress', methods=['POST'])
def process_compress():
    if 'pdf_file' not in request.files:
        return redirect(url_for('index', error_msg="No file part", active_tab="compress"))
        
    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return redirect(url_for('index', error_msg="No selected file", active_tab="compress"))
        
    if pdf_file and pdf_file.filename.lower().endswith('.pdf'):
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        out_filename = "COMPRESSED_" + pdf_file.filename
        out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
        
        pdf_file.save(pdf_path)
        
        try:
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            
            for page in reader.pages:
                new_page = writer.add_page(page)
                new_page.compress_content_streams()
                
            for page in writer.pages:
                for img in page.images:
                    img.replace(img.image, quality=60)
                
            with open(out_path, "wb") as f:
                writer.write(f)
                
            orig_size = os.path.getsize(pdf_path)
            new_size = os.path.getsize(out_path)
            
            try:
                os.remove(pdf_path)
            except:
                pass
                
            reduction = ((orig_size - new_size) / orig_size) * 100
            success_msg = f"Successfully compressed! Size reduced from {orig_size/1024:.1f} KB to {new_size/1024:.1f} KB ({reduction:.1f}% reduction)."
            return redirect(url_for('index', success_msg=success_msg, result_file=out_filename, active_tab="compress"))
        except Exception as e:
            try:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
            except:
                pass
            return redirect(url_for('index', error_msg=f"Error occurred: {str(e)}", active_tab="compress"))
            
    return redirect(url_for('index', error_msg="Invalid file format. Please upload a PDF.", active_tab="compress"))

@app.route('/process_convert', methods=['POST'])
def process_convert():
    if 'pdf_file' not in request.files:
        return redirect(url_for('index', error_msg="No file part", active_tab="pdf2word"))
        
    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return redirect(url_for('index', error_msg="No selected file", active_tab="pdf2word"))
        
    if pdf_file and pdf_file.filename.lower().endswith('.pdf'):
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        docx_filename = pdf_file.filename[:-4] + ".docx"
        docx_path = os.path.join(app.config['UPLOAD_FOLDER'], docx_filename)
        
        pdf_file.save(pdf_path)
        
        try:
            from pdf2docx import Converter
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            
            try:
                os.remove(pdf_path)
            except:
                pass
                
            return redirect(url_for('index', success_msg="Conversion successful!", result_file=docx_filename, active_tab="pdf2word"))
        except Exception as e:
            try:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
            except:
                pass
            return redirect(url_for('index', error_msg=f"Error occurred during conversion: {str(e)}", active_tab="pdf2word"))
            
    return redirect(url_for('index', error_msg="Invalid file format. Please upload a PDF.", active_tab="pdf2word"))

@app.route('/process_word2pdf', methods=['POST'])
def process_word2pdf():
    if 'doc_file' not in request.files:
        return redirect(url_for('index', error_msg="No file part", active_tab="word2pdf"))
        
    doc_file = request.files['doc_file']
    if doc_file.filename == '':
        return redirect(url_for('index', error_msg="No selected file", active_tab="word2pdf"))
        
    if doc_file and doc_file.filename.lower().endswith('.docx'):
        docx_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_file.filename)
        pdf_filename = doc_file.filename[:-5] + ".pdf"
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
        
        doc_file.save(docx_path)
        
        try:
            import pythoncom
            pythoncom.CoInitialize()
            
            import win32com.client
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            try:
                doc = word.Documents.Open(docx_path)
                doc.SaveAs(pdf_path, FileFormat=17) # 17 is wdFormatPDF
                doc.Close()
            finally:
                word.Quit()
                
            try:
                os.remove(docx_path)
            except:
                pass
                
            return redirect(url_for('index', success_msg="Conversion successful!", result_file=pdf_filename, active_tab="word2pdf"))
        except Exception as e:
            try:
                if os.path.exists(docx_path):
                    os.remove(docx_path)
            except:
                pass
            return redirect(url_for('index', error_msg=f"Error occurred during conversion: {str(e)}", active_tab="word2pdf"))
            
    return redirect(url_for('index', error_msg="Invalid file format. Please upload a Word Document (.docx).", active_tab="word2pdf"))

@app.route('/process_pdf_solver', methods=['POST'])
def process_pdf_solver():
    engine = request.form.get('engine', 'gemini')
    api_key = request.form.get('api_key')
    
    if engine != 'local' and not api_key:
        return redirect(url_for('index', error_msg="API Key is required for AI modes.", active_tab="pdf_solver"))
        
    try:
        pdf_name, pdf_path, is_temp_pdf = resolve_file('server_pdf_file', 'pdf_file', '.pdf')
    except Exception as e:
        return redirect(url_for('index', error_msg=str(e), api_key_val=api_key, active_tab="pdf_solver"))
        
    sanitized_path = os.path.join(app.config['UPLOAD_FOLDER'], "SANITIZED_" + pdf_name)
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        with open(sanitized_path, "wb") as f:
            writer.write(f)
        os.remove(pdf_path)
        pdf_path = sanitized_path
    except Exception as clean_err:
        print(f"Sanitization failed: {clean_err}")
        if os.path.exists(sanitized_path):
            try: os.remove(sanitized_path)
            except: pass
            
    try:
        response_text = None
        last_err = None
        
        prompt = (
            "Berikut adalah dokumen PDF Turnitin. Silakan analisis HANYA pada bagian teks yang memiliki warna highlight/sorotan plagiasi. "
            "Abaikan teks yang bersih (tanpa warna). Hasilkan output berupa daftar teks asli dan hasil parafrasenya sesuai aturan format Markdown yang telah ditetapkan."
        )
        
        if engine == 'local':
            reader = PdfReader(pdf_path)
            full_text = ""
            for idx, page in enumerate(reader.pages):
                text = page.extract_text() or ""
                if text.strip():
                    full_text += f"\n### Halaman {idx+1}\n"
                    for line in text.split('\n'):
                        if line.strip():
                            para = offline_paraphrase(line.strip())
                            full_text += f"* **Teks Asli:** \"{line.strip()}\"\n* **Hasil Parafrase:** \"**{para}**\"\n\n"
            response_text = full_text
            
        elif engine == 'openrouter':
            import requests
            import base64
            with open(pdf_path, "rb") as f:
                pdf_data = base64.b64encode(f.read()).decode("utf-8")
            
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": "openrouter/free",
                "messages": [
                    {"role": "system", "content": SYSTEM_INSTRUCTION_PDF_SOLVER},
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:application/pdf;base64,{pdf_data}"
                                }
                            }
                        ]
                    }
                ],
                "max_tokens": 3000
            }
            res = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)
            if res.status_code == 200:
                response_text = res.json()["choices"][0]["message"]["content"]
            else:
                last_err = res.text
        else:
            genai.configure(api_key=api_key)
            uploaded_file = genai.upload_file(path=pdf_path, mime_type='application/pdf')
            
            import time
            for _ in range(30):
                if uploaded_file.state.name == "ACTIVE":
                    break
                elif uploaded_file.state.name == "FAILED":
                    raise ValueError("Gagal memproses file PDF di server Google API.")
                time.sleep(2)
                uploaded_file = genai.get_file(uploaded_file.name)
                
            models_to_try = [
                'gemini-2.0-flash-latest', 'gemini-2.0-flash', 
                'gemini-2.5-flash', 'gemini-2.5-flash-lite', 
                'gemini-1.5-flash', 'gemini-1.5-flash-latest', 
                'gemini-1.5-pro', 'gemini-2.5-pro'
            ]
            
            for model_name in models_to_try:
                try:
                    model = genai.GenerativeModel(model_name, system_instruction=SYSTEM_INSTRUCTION_PDF_SOLVER)
                    response = model.generate_content([uploaded_file, prompt])
                    if response and response.text:
                        response_text = response.text
                        break
                except Exception as model_err:
                    print(f"Failed using model {model_name}: {model_err}")
                    last_err = model_err
                    
            try:
                genai.delete_file(uploaded_file.name)
            except:
                pass
            
        if not response_text:
            raise ValueError(f"Paraphrase engine returned empty response or failed. Last error: {last_err}")
            
        # Parse replacements from response_text and save to session json file
        import json
        try:
            reps = []
            current_original = None
            current_paraphrase = None
            for line in response_text.split('\n'):
                txt = line.strip()
                if not txt:
                    continue
                orig_match = re.search(r'\*\*Teks Asli:\*\*\s*(.*)$', txt)
                para_match = re.search(r'\*\*Hasil Parafrase:\*\*\s*(.*)$', txt)
                if orig_match:
                    val = orig_match.group(1).strip()
                    if val.startswith('"') and val.endswith('"'): val = val[1:-1]
                    current_original = val.strip()
                elif para_match:
                    val = para_match.group(1).strip()
                    if val.startswith('**'): val = val[2:]
                    if val.endswith('**'): val = val[:-2]
                    val = val.strip()
                    if val.startswith('"') and val.endswith('"'): val = val[1:-1]
                    current_paraphrase = val.strip()
                    if current_original:
                        current_paraphrase = current_paraphrase.replace("**", "")
                        reps.append([current_original, current_paraphrase])
                        current_original = None
                        current_paraphrase = None
            
            if reps:
                active_reps_path = os.path.join(app.config['UPLOAD_FOLDER'], 'active_replacements.json')
                with open(active_reps_path, 'w') as f:
                    json.dump(reps, f)
        except Exception as parse_save_err:
            print(f"Failed to save active replacements: {parse_save_err}")
            
        out_filename = "turnitin_slayer_result.docx"
        out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
        
        doc = Document()
        paragraphs = response_text.split('\n')
        for para in paragraphs:
            if para.strip():
                doc.add_paragraph(para.strip())
        doc.save(out_path)
        
        if is_temp_pdf:
            try:
                os.remove(pdf_path)
            except:
                pass
            
        return redirect(url_for('index', success_msg="PDF successfully analyzed! Paraphrases generated successfully.", result_file=out_filename, api_key_val=api_key, active_tab="pdf_solver"))
        
    except Exception as e:
        try:
            if is_temp_pdf and os.path.exists(pdf_path):
                os.remove(pdf_path)
        except:
            pass
        return redirect(url_for('index', error_msg=f"Error occurred: {str(e)}", api_key_val=api_key, active_tab="pdf_solver"))

COMPARE_RESULT_TEMPLATE = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Perbandingan Dokumen - Slayer Suite</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-color: #fcfcfc;
            --card-bg: #ffffff;
            --text-color: #1a1a1a;
            --secondary-text: #666666;
            --accent-color: #000000;
            --border-color: #e5e5e5;
            --red-highlight: #fde8e8;
            --green-highlight: #eafaf1;
            --red-text: #9b1c1c;
            --green-text: #0f5132;
        }

        body {
            font-family: 'Inter', sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
            margin: 0;
            padding: 20px;
            font-size: 14px;
        }

        .header {
            max-width: 1200px;
            margin: 0 auto 20px auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 20px;
        }

        .header h1 {
            font-size: 1.5rem;
            margin: 0;
            font-weight: 700;
        }

        .header p {
            margin: 5px 0 0 0;
            color: var(--secondary-text);
            font-size: 0.85rem;
        }

        .btn-back {
            display: inline-block;
            padding: 8px 16px;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            background: #ffffff;
            color: var(--text-color);
            font-weight: 500;
            text-decoration: none;
            transition: background 0.15s;
            cursor: pointer;
        }

        .btn-back:hover {
            background: #f5f5f4;
        }

        .comparison-container {
            max-width: 1200px;
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            gap: 15px;
        }

        .grid-header {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            font-weight: 600;
            font-size: 0.9rem;
            padding: 10px 0;
            border-bottom: 2px solid var(--border-color);
        }

        .row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 15px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        }

        .col {
            padding: 10px;
            border-radius: 6px;
            line-height: 1.6;
            white-space: pre-wrap;
            word-break: break-word;
        }

        .col-orig {
            background: var(--red-highlight);
            color: var(--red-text);
            border-left: 4px solid #f8b4b4;
        }

        .col-para {
            background: var(--green-highlight);
            color: var(--green-text);
            border-left: 4px solid #84e1bc;
        }

        .para-num {
            font-size: 0.75rem;
            color: var(--secondary-text);
            font-weight: 600;
            margin-bottom: 5px;
            text-transform: uppercase;
        }

        .no-diff {
            text-align: center;
            padding: 50px;
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            font-size: 1.1rem;
            color: var(--secondary-text);
        }
    </style>
</head>
<body>
    <div class="header">
        <div>
            <h1>Hasil Perbandingan Dokumen</h1>
            <p>Membandingkan <strong>{{ orig_name }}</strong> (Kiri) dan <strong>{{ para_name }}</strong> (Kanan)</p>
        </div>
        <a href="/" class="btn-back">⬅️ Kembali ke Dashboard</a>
    </div>

    <div class="comparison-container">
        {% if diffs %}
            <div class="grid-header">
                <div>Original Text (Skripsi Asli)</div>
                <div>Paraphrased Text (Hasil Parafrase)</div>
            </div>
            
            {% for diff in diffs %}
            <div class="row">
                <div>
                    <div class="para-num">Paragraf {{ diff.paragraph_num }}</div>
                    <div class="col col-orig">{{ diff.original }}</div>
                </div>
                <div>
                    <div class="para-num">Paragraf {{ diff.paragraph_num }}</div>
                    <div class="col col-para">{{ diff.paraphrased }}</div>
                </div>
            </div>
            {% endfor %}
        {% else %}
            <div class="no-diff">
                🎉 Dokumen identik! Tidak ditemukan adanya perbedaan kata/kalimat di antara kedua file tersebut.
            </div>
        {% endif %}
    </div>
</body>
</html>
"""

@app.route('/process_compare_docs', methods=['POST'])
def process_compare_docs():
    try:
        orig_name, orig_path, is_temp_orig = resolve_file('server_original_doc', 'original_doc', '.docx')
    except Exception as e:
        return redirect(url_for('index', error_msg=str(e), active_tab="compare_docs"))
        
    try:
        para_name, para_path, is_temp_para = resolve_file('server_paraphrased_doc', 'paraphrased_doc', '.docx')
    except Exception as e:
        if is_temp_orig:
            try: os.remove(orig_path)
            except: pass
        return redirect(url_for('index', error_msg=str(e), active_tab="compare_docs"))
        
    try:
        doc_orig = Document(orig_path)
        doc_para = Document(para_path)
        
        diffs = []
        for idx, (p_orig, p_para) in enumerate(zip(doc_orig.paragraphs, doc_para.paragraphs)):
            txt_orig = p_orig.text.strip()
            txt_para = p_para.text.strip()
            if txt_orig != txt_para:
                diffs.append({
                    'paragraph_num': idx + 1,
                    'original': txt_orig,
                    'paraphrased': txt_para
                })
                
        if is_temp_orig:
            try: os.remove(orig_path)
            except: pass
        if is_temp_para:
            try: os.remove(para_path)
            except: pass
            
        return render_template_string(COMPARE_RESULT_TEMPLATE, diffs=diffs, orig_name=orig_name, para_name=para_name)
    except Exception as e:
        if is_temp_orig:
            try: os.remove(orig_path)
            except: pass
        if is_temp_para:
            try: os.remove(para_path)
            except: pass
        return redirect(url_for('index', error_msg=f"Error occurred during comparison: {str(e)}", active_tab="compare_docs"))

@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return redirect(url_for('index', error_msg="File not found."))

if __name__ == '__main__':
    app.run(debug=True, port=5000)
