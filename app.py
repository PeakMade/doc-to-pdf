from flask import Flask, request, send_file, render_template, flash, redirect, url_for
import os
from werkzeug.utils import secure_filename
import tempfile
from pathlib import Path
import platform
import subprocess

# Use docx2pdf for Windows (local dev) and LibreOffice for Docker/Linux
IS_WINDOWS = platform.system() == 'Windows'

if IS_WINDOWS:
    try:
        from docx2pdf import convert as docx2pdf_convert
        import pythoncom
        USE_DOCX2PDF = True
    except ImportError:
        USE_DOCX2PDF = False
else:
    USE_DOCX2PDF = False

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this in production
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_docx_to_pdf():
    # Check if file was uploaded
    if 'file' not in request.files:
        flash('No file uploaded')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        # Secure the filename
        filename = secure_filename(file.filename)
        
        # Create temporary paths
        docx_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        pdf_filename = Path(filename).stem + '.pdf'
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
        
        try:
            # Save the uploaded file
            file.save(docx_path)
            
            # Convert to PDF based on platform
            if IS_WINDOWS and USE_DOCX2PDF:
                # Windows: Use docx2pdf with COM
                pythoncom.CoInitialize()
                try:
                    docx2pdf_convert(docx_path, pdf_path)
                finally:
                    pythoncom.CoUninitialize()
            else:
                # Docker/Linux: Use LibreOffice for high-fidelity conversion
                # LibreOffice preserves all formatting, images, tables, and graphics
                result = subprocess.run([
                    'libreoffice',
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', app.config['UPLOAD_FOLDER'],
                    docx_path
                ], capture_output=True, text=True, timeout=120)
                
                if result.returncode != 0:
                    raise Exception(f'LibreOffice conversion failed: {result.stderr}')
            
            # Send the PDF file
            return send_file(
                pdf_path,
                as_attachment=True,
                download_name=pdf_filename,
                mimetype='application/pdf'
            )
        
        except Exception as e:
            flash(f'Error converting file: {str(e)}')
            return redirect(url_for('index'))
        
        finally:
            # Clean up temporary files
            try:
                if os.path.exists(docx_path):
                    os.remove(docx_path)
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
            except:
                pass
    
    else:
        flash('Invalid file type. Please upload a .docx file')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # Get port from environment variable for Azure, default to 5005 for local dev
    port = int(os.environ.get('PORT', 5005))
    app.run(debug=False, host='0.0.0.0', port=port)
