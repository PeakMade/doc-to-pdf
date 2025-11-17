from flask import Flask, request, send_file, render_template, flash, redirect, url_for, jsonify
import os
from werkzeug.utils import secure_filename
import tempfile
from pathlib import Path
import platform
import subprocess
import io

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

def perform_conversion(docx_path, pdf_path):
    """
    Core conversion logic shared by both web UI and API endpoints.
    
    Args:
        docx_path: Path to input DOCX file
        pdf_path: Path where PDF should be created
    
    Raises:
        Exception: If conversion fails
    """
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
            '--outdir', os.path.dirname(pdf_path),
            docx_path
        ], capture_output=True, text=True, timeout=120)
        
        if result.returncode != 0:
            raise Exception(f'LibreOffice conversion failed: {result.stderr}')
        
        # LibreOffice creates PDF with same base name as input
        expected_pdf = os.path.join(
            os.path.dirname(pdf_path),
            Path(os.path.basename(docx_path)).stem + '.pdf'
        )
        if expected_pdf != pdf_path and os.path.exists(expected_pdf):
            os.rename(expected_pdf, pdf_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/health', methods=['GET'])
def health_check():
    """
    Health check endpoint for Azure Container Apps monitoring.
    Returns 200 if service is healthy.
    """
    return jsonify({
        'status': 'healthy',
        'service': 'docx-to-pdf-converter',
        'timestamp': __import__('datetime').datetime.utcnow().isoformat()
    }), 200

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
            
            # Convert using shared logic
            perform_conversion(docx_path, pdf_path)
            
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

@app.route('/api/convert', methods=['POST'])
def api_convert_docx_to_pdf():
    """
    API endpoint for programmatic DOCX to PDF conversion.
    Returns PDF file directly or JSON error response.
    """
    # Optional: API key authentication (uncomment to enable)
    # api_key = request.headers.get('X-API-Key')
    # if api_key != os.environ.get('API_KEY'):
    #     return jsonify({'error': 'Unauthorized'}), 401
    
    # Validate request
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only .docx files are accepted'}), 400
    
    # Create unique temporary paths
    import uuid
    filename = secure_filename(file.filename)
    unique_id = str(uuid.uuid4())
    docx_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{filename}")
    pdf_filename = Path(filename).stem + '.pdf'
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{pdf_filename}")
    
    try:
        # Save and convert
        file.save(docx_path)
        perform_conversion(docx_path, pdf_path)
        
        # Read PDF into memory for return
        with open(pdf_path, 'rb') as pdf_file:
            pdf_data = io.BytesIO(pdf_file.read())
            pdf_data.seek(0)
        
        return send_file(
            pdf_data,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=pdf_filename
        )
    
    except subprocess.TimeoutExpired:
        return jsonify({'error': 'Conversion timeout - file may be too large or complex'}), 500
    except Exception as e:
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500
    finally:
        # Always cleanup temporary files
        for path in [docx_path, pdf_path]:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except:
                pass

if __name__ == '__main__':
    # Get port from environment variable for Azure, default to 5005 for local dev
    port = int(os.environ.get('PORT', 5005))
    app.run(debug=False, host='0.0.0.0', port=port)
