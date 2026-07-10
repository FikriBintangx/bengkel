import os
from flask import Flask, request, render_template_string, send_file, redirect, url_for
from pdf2docx import Converter

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Notion-style Black and White HTML layout for PDF to Word
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF to Word Converter</title>
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
            max-width: 600px;
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

        .file-upload-wrapper {
            position: relative;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 30px;
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
        }

        footer {
            margin-top: 50px;
            font-size: 0.8rem;
            color: var(--secondary-text);
            text-align: center;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>PDF to Word</h1>
            <p class="subtitle">Convert PDF documents to editable Docx files instantly.</p>
        </div>

        <div class="divider"></div>

        {% if success_msg %}
            <div class="alert-box">{{ success_msg }}</div>
        {% endif %}

        {% if error_msg %}
            <div class="error-box">{{ error_msg }}</div>
        {% endif %}

        <form action="/convert" method="POST" enctype="multipart/form-data">
            <div class="form-group">
                <label>Select PDF File</label>
                <div class="file-upload-wrapper">
                    <span class="upload-icon">📕</span>
                    <div class="file-name-label" id="pdf-label">Choose a file or drag it here</div>
                    <input type="file" name="pdf_file" accept=".pdf" required onchange="updateLabel(this)">
                </div>
            </div>

            <button type="submit" class="btn-submit">Convert to Word</button>
        </form>
        
        {% if result_file %}
            <div style="margin-top: 25px; text-align: center;">
                <a href="/download/{{ result_file }}" class="btn-download">📥 Download Word Document</a>
            </div>
        {% endif %}

        <footer>Powered by Antigravity AI & pdf2docx</footer>
    </div>

    <script>
        function updateLabel(input) {
            const label = document.getElementById('pdf-label');
            if (input.files && input.files.length > 0) {
                label.innerText = input.files[0].name;
                label.style.color = 'var(--text-color)';
            } else {
                label.innerText = 'Choose a file or drag it here';
                label.style.color = 'var(--secondary-text)';
            }
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    success_msg = request.args.get('success_msg')
    error_msg = request.args.get('error_msg')
    result_file = request.args.get('result_file')
    return render_template_string(HTML_TEMPLATE, success_msg=success_msg, error_msg=error_msg, result_file=result_file)

@app.route('/convert', methods=['POST'])
def convert():
    if 'pdf_file' not in request.files:
        return redirect(url_for('index', error_msg="No file part"))
        
    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return redirect(url_for('index', error_msg="No selected file"))
        
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
                
            return redirect(url_for('index', success_msg="Conversion successful!", result_file=docx_filename))
        except Exception as e:
            return redirect(url_for('index', error_msg=f"Error occurred during conversion: {str(e)}"))
            
    return redirect(url_for('index', error_msg="Invalid file format. Please upload a PDF."))

@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return redirect(url_for('index', error_msg="File not found."))

if __name__ == '__main__':
    app.run(debug=True, port=5002)
