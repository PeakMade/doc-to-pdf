# DOCX to PDF Converter - Flask App

A simple Flask web application that converts Microsoft Word (.docx) files to PDF format.

## Features

- Clean, modern web interface
- Drag-and-drop file upload support
- Instant conversion from DOCX to PDF
- Automatic file download after conversion
- File size limit: 16MB
- Temporary file cleanup

## Installation

1. Install the required dependencies:
```powershell
pip install -r requirements.txt
```

**Note for Windows users:** The `docx2pdf` library requires Microsoft Word to be installed on Windows systems, as it uses COM automation to perform the conversion.

## Usage

1. Run the Flask application:
```powershell
python app.py
```

2. Open your browser and navigate to:
```
http://127.0.0.1:5000
```

3. Upload a .docx file using the web interface

4. The converted PDF will automatically download to your browser

## Alternative for Linux/Mac

If you're on Linux or Mac (or Windows without MS Word), you can modify the code to use an alternative conversion library like `pypandoc` or `libreoffice`:

### Using LibreOffice (cross-platform):
```python
import subprocess

def convert_docx_to_pdf_libreoffice(docx_path, output_dir):
    subprocess.run([
        'soffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', output_dir,
        docx_path
    ])
```

## Security Considerations for Production

- Change the `app.secret_key` to a secure random value
- Implement proper authentication if needed
- Add rate limiting to prevent abuse
- Use environment variables for configuration
- Deploy with a production WSGI server (gunicorn, waitress)
- Add virus scanning for uploaded files

## License

MIT
