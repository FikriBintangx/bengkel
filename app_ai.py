import os
import re
import time
import docx
from docx import Document
from flask import Flask, request, render_template_string, send_file, redirect, url_for, jsonify
import google.generativeai as genai

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Notion-style Black and White HTML layout for AI Pro version
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Turnitin Slayer AI Pro</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-color: #ffffff;
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
            padding: 80px 20px;
            box-sizing: border-box;
        }

        .header {
            margin-bottom: 40px;
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

        .divider {
            height: 1px;
            background-color: var(--border-color);
            margin: 30px 0;
        }

        .form-group {
            margin-bottom: 28px;
        }

        label {
            display: block;
            font-weight: 600;
            margin-bottom: 8px;
            font-size: 0.9rem;
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
            padding: 10px 20px;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            background: #ffffff;
            color: var(--text-color);
            font-weight: 600;
            font-size: 0.95rem;
            text-decoration: none;
            transition: background 0.15s ease;
            box-shadow: 0 1px 2px rgba(15, 15, 15, 0.05);
        }

        .btn-download:hover {
            background: var(--hover-bg);
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

        footer {
            margin-top: 50px;
            font-size: 0.8rem;
            color: var(--secondary-text);
            text-align: center;
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
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Turnitin Slayer AI Pro</h1>
            <p class="subtitle">Bypass Turnitin Detection with Advanced Gemini AI Paraphrasing Rules.</p>
        </div>

        <div class="divider"></div>

        {% if success_msg %}
            <div class="alert-box">{{ success_msg }}</div>
        {% endif %}

        {% if error_msg %}
            <div class="error-box">{{ error_msg }}</div>
        {% endif %}

        <form action="/process" method="POST" enctype="multipart/form-data">
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
                    <div class="file-name-label" id="orig-label">Choose a file or drag it here</div>
                    <input type="file" name="original_doc" accept=".docx" required onchange="updateLabel(this, 'orig-label')">
                </div>
            </div>
            
            <div class="form-group">
                <label>Paraphrase Rules (Active & Updated)</label>
                <div class="rules-container">
                    <ul>
                        <li><strong>Aturan Umum:</strong> Jangan mengubah makna utama, pertahankan fakta/angka/statistik/hasil penelitian. Jangan menambah info baru, mengurangi info penting, membuat opini/asumsi baru, dan pertahankan tujuan penulis.</li>
                        <li><strong>Struktur Kalimat:</strong> Variasikan struktur, hindari pengulangan pola, pecah kalimat panjang, gabungkan kalimat pendek, variasikan posisi S-P-O-K.</li>
                        <li><strong>Diksi & Gaya Akademik:</strong> Gunakan sinonim kontekstual yang presisi, istilah akademik baku (non-percakapan), bahasa formal objektif netral, konsisten, pertahankan nada akademik.</li>
                        <li><strong>Paragraf & Istilah:</strong> Koheren, transisi alami, variasi pembuka/penutup, hindari kalimat tunggal jika tidak perlu. Jangan ubah istilah teknis, nama teori/metode/variabel, rumus/simbol/satuan.</li>
                        <li><strong>Sitasi & Kualitas:</strong> Pertahankan nama penulis/tahun/format sitasi, DOI, URL, daftar pustaka. Lebih jelas, ringkas, hilangkan pengulangan ide, ejaan sesuai PUEBI, hasil terdengar natural.</li>
                    </ul>
                </div>
            </div>

            <button type="submit" class="btn-submit">Slay Turnitin Now</button>
        </form>
        
        {% if result_file %}
            <div style="margin-top: 25px; text-align: center;">
                <a href="/download/{{ result_file }}" class="btn-download">📥 Download Paraphrased Docx</a>
            </div>
        {% endif %}

        <footer>Powered by Antigravity AI & Gemini API</footer>
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

        function checkAPIKey() {
            const apiKey = document.getElementById('api_key_input').value.strip ? document.getElementById('api_key_input').value.strip() : document.getElementById('api_key_input').value;
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

@app.route('/process', methods=['POST'])
def process():
    api_key = request.form.get('api_key')
    if not api_key:
        return redirect(url_for('index', error_msg="Gemini API Key is required."))
        
    if 'original_doc' not in request.files:
        return redirect(url_for('index', error_msg="Please select a document.", api_key_val=api_key))
        
    orig_file = request.files['original_doc']
    if orig_file.filename == '':
        return redirect(url_for('index', error_msg="Invalid file name.", api_key_val=api_key))
        
    if not orig_file.filename.lower().endswith('.docx'):
        return redirect(url_for('index', error_msg="Format file harus berupa dokumen Word (.docx). File PDF atau format lain tidak didukung.", api_key_val=api_key))

    orig_path = os.path.join(app.config['UPLOAD_FOLDER'], orig_file.filename)
    orig_file.save(orig_path)
    
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash', system_instruction=PARAPHRASE_SYSTEM_PROMPT)
        
        try:
            doc = Document(orig_path)
        except Exception as docx_err:
            raise Exception(f"File dokumen tidak bisa dibuka atau rusak. Pastikan file bertipe .docx (bukan PDF yang direname menjadi .docx). Detail error: {docx_err}")

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
            
        return redirect(url_for('index', success_msg=f"Successfully paraphrased {replaced_count} paragraphs!", result_file=out_filename, api_key_val=api_key))
        
    except Exception as e:
        try:
            if os.path.exists(orig_path):
                os.remove(orig_path)
        except:
            pass
        return redirect(url_for('index', error_msg=f"Error occurred: {str(e)}", api_key_val=api_key))

@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return redirect(url_for('index', error_msg="Result file not found."))

if __name__ == '__main__':
    app.run(debug=True, port=5001)
