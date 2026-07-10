import os
import re
import time
from docx import Document
from flask import Flask, request, render_template_string, send_file, redirect, url_for, jsonify
import google.generativeai as genai
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Paraphrase System Prompt for Gemini
PARAPHRASE_SYSTEM_PROMPT = """Anda adalah asisten akademik profesional yang berspesialisasi dalam menulis ulang karya ilmiah (skripsi/tesis) untuk menurunkan persentase kemiripan (Turnitin) di tingkat universitas.

Tugas Anda adalah memparafrasekan teks bahasa Indonesia yang diberikan oleh pengguna dengan mengikuti instruksi ketat berikut:

A. Aturan Umum
- Jangan mengubah makna utama.
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
    </style>
</head>
<body>
    <div class="container">
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
                    <div class="form-group">
                        <label>Original Skripsi Document (.docx)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">📄</span>
                            <div class="file-name-label" id="manual-orig-label">Choose a file or drag it here</div>
                            <input type="file" name="original_doc" accept=".docx" required onchange="updateLabel(this, 'manual-orig-label')">
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Paraphrase Reference Document (.docx)</label>
                        <div class="file-upload-wrapper">
                            <span class="upload-icon">✍️</span>
                            <div class="file-name-label" id="manual-ref-label">Choose a file or drag it here</div>
                            <input type="file" name="reference_doc" accept=".docx" required onchange="updateLabel(this, 'manual-ref-label')">
                        </div>
                    </div>

                    <button type="submit" class="btn-submit">Process & Replace</button>
                </form>
            </div>

            <!-- TAB 2: TURNITIN AI PRO -->
            <div id="ai" class="tab-content">
                <form action="/process_ai" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Gemini API Key</label>
                        <div class="api-input-container">
                            <input type="text" id="api_key_input" name="api_key" placeholder="Enter your Gemini API Key here" required value="{{ api_key_val }}">
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

            <footer>Powered by Antigravity AI</footer>
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
    return render_template_string(HTML_TEMPLATE, success_msg=success_msg, error_msg=error_msg, result_file=result_file, api_key_val=api_key_val)

@app.route('/check_api', methods=['POST'])
def check_api():
    data = request.get_json()
    api_key = data.get('api_key')
    if not api_key:
        return jsonify({'status': 'error', 'message': 'API Key kosong'})
        
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content("Tes koneksi. Balas dengan kata OK saja.")
        if response.text:
            return jsonify({'status': 'success'})
        else:
            return jsonify({'status': 'error', 'message': 'Respons kosong dari API'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/process_manual', methods=['POST'])
def process_manual():
    if 'original_doc' not in request.files or 'reference_doc' not in request.files:
        return redirect(url_for('index', error_msg="Please select both documents.", active_tab="manual"))
        
    orig_file = request.files['original_doc']
    ref_file = request.files['reference_doc']
    
    if orig_file.filename == '' or ref_file.filename == '':
        return redirect(url_for('index', error_msg="Invalid file name.", active_tab="manual"))
        
    if orig_file and ref_file:
        orig_path = os.path.join(app.config['UPLOAD_FOLDER'], orig_file.filename)
        ref_path = os.path.join(app.config['UPLOAD_FOLDER'], ref_file.filename)
        
        orig_file.save(orig_path)
        ref_file.save(ref_path)
        
        try:
            reps = parse_revisi(ref_path)
            if not reps:
                return redirect(url_for('index', error_msg="No replacements found in reference doc.", active_tab="manual"))
                
            out_filename = "PARAFRASED_" + orig_file.filename
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            
            replaced_count = do_smart_replacements(orig_path, reps, out_path)
            
            try:
                os.remove(orig_path)
                os.remove(ref_path)
            except:
                pass
                
            return redirect(url_for('index', success_msg=f"Success! Replaced {replaced_count} matches.", result_file=out_filename, active_tab="manual"))
        except Exception as e:
            return redirect(url_for('index', error_msg=f"Error occurred: {str(e)}", active_tab="manual"))
            
    return redirect(url_for('index', error_msg="Unknown error.", active_tab="manual"))

@app.route('/process_ai', methods=['POST'])
def process_ai():
    api_key = request.form.get('api_key')
    if not api_key:
        return redirect(url_for('index', error_msg="Gemini API Key is required.", active_tab="ai"))
        
    if 'original_doc' not in request.files:
        return redirect(url_for('index', error_msg="Please select a document.", api_key_val=api_key, active_tab="ai"))
        
    orig_file = request.files['original_doc']
    if orig_file.filename == '':
        return redirect(url_for('index', error_msg="Invalid file name.", api_key_val=api_key, active_tab="ai"))
        
    if not orig_file.filename.lower().endswith('.docx'):
        return redirect(url_for('index', error_msg="Format file harus berupa dokumen Word (.docx).", api_key_val=api_key, active_tab="ai"))

    orig_path = os.path.join(app.config['UPLOAD_FOLDER'], orig_file.filename)
    orig_file.save(orig_path)
    
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash', system_instruction=PARAPHRASE_SYSTEM_PROMPT)
        
        try:
            doc = Document(orig_path)
        except Exception as docx_err:
            raise Exception(f"File dokumen tidak bisa dibuka atau rusak: {docx_err}")

        replaced_count = 0
        
        for p in doc.paragraphs:
            original_text = p.text.strip()
            if len(original_text) > 20 and not original_text.startswith("BAB ") and not original_text.startswith("DAFTAR PUSTAKA") and not original_text.isupper():
                try:
                    time.sleep(4.0) # Delay to avoid 429 Rate Limit
                    response = model.generate_content(original_text)
                    para_text = response.text.strip()
                    
                    if para_text:
                        if len(p.runs) > 0:
                            p.runs[0].text = para_text
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.text = para_text
                        replaced_count += 1
                except Exception as api_err:
                    print(f"API Error at paragraph: {original_text[:50]}... Error: {api_err}")
                    
        out_filename = "AI_PARAFRASED_" + orig_file.filename
        out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
        doc.save(out_path)
        
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

@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return redirect(url_for('index', error_msg="File not found."))

if __name__ == '__main__':
    app.run(debug=True, port=5000)
