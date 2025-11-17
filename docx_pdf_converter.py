"""
DOCX to PDF Converter Module

A reusable module for converting Microsoft Word (.docx) files to PDF format.
Works cross-platform with automatic detection:
- Windows: Uses docx2pdf with COM automation (requires MS Word installed)
- Linux/Docker: Uses LibreOffice headless mode

Usage:
    from docx_pdf_converter import convert_docx_to_pdf
    
    pdf_path = convert_docx_to_pdf(
        docx_path='/path/to/document.docx',
        output_dir='/path/to/output'  # Optional, defaults to same dir as input
    )
"""

import os
import platform
import subprocess
from pathlib import Path

# Detect platform and available conversion methods
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


def convert_docx_to_pdf(docx_path, output_dir=None):
    """
    Convert a DOCX file to PDF format.
    
    Args:
        docx_path (str): Path to the input .docx file
        output_dir (str, optional): Directory where the PDF should be saved.
                                   If None, saves in the same directory as the input file.
    
    Returns:
        str: Path to the generated PDF file
    
    Raises:
        FileNotFoundError: If the input docx file doesn't exist
        Exception: If the conversion fails
    
    Examples:
        >>> pdf_path = convert_docx_to_pdf('/tmp/document.docx')
        >>> print(pdf_path)
        /tmp/document.pdf
        
        >>> pdf_path = convert_docx_to_pdf('/tmp/document.docx', output_dir='/tmp/pdfs')
        >>> print(pdf_path)
        /tmp/pdfs/document.pdf
    """
    # Validate input file exists
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Input file not found: {docx_path}")
    
    # Determine output directory
    if output_dir is None:
        output_dir = os.path.dirname(docx_path)
    else:
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
    
    # Generate output filename
    input_filename = os.path.basename(docx_path)
    pdf_filename = Path(input_filename).stem + '.pdf'
    pdf_path = os.path.join(output_dir, pdf_filename)
    
    # Convert based on platform
    if IS_WINDOWS and USE_DOCX2PDF:
        # Windows: Use docx2pdf with COM automation
        pythoncom.CoInitialize()
        try:
            docx2pdf_convert(docx_path, pdf_path)
        finally:
            pythoncom.CoUninitialize()
    else:
        # Linux/Docker: Use LibreOffice for high-fidelity conversion
        # LibreOffice preserves all formatting, images, tables, and graphics
        result = subprocess.run([
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            docx_path
        ], capture_output=True, text=True, timeout=120)
        
        if result.returncode != 0:
            raise Exception(f'LibreOffice conversion failed: {result.stderr}')
    
    # Verify the PDF was created
    if not os.path.exists(pdf_path):
        raise Exception(f'PDF file was not created: {pdf_path}')
    
    return pdf_path


def convert_docx_to_pdf_bytes(docx_path):
    """
    Convert a DOCX file to PDF and return the PDF content as bytes.
    Useful for in-memory operations or direct HTTP responses.
    
    Args:
        docx_path (str): Path to the input .docx file
    
    Returns:
        bytes: The PDF file content as bytes
    
    Raises:
        FileNotFoundError: If the input docx file doesn't exist
        Exception: If the conversion fails
    
    Examples:
        >>> pdf_bytes = convert_docx_to_pdf_bytes('/tmp/document.docx')
        >>> with open('/tmp/output.pdf', 'wb') as f:
        ...     f.write(pdf_bytes)
    """
    import tempfile
    
    # Create a temporary directory for conversion
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = convert_docx_to_pdf(docx_path, output_dir=temp_dir)
        
        # Read the PDF content
        with open(pdf_path, 'rb') as pdf_file:
            pdf_bytes = pdf_file.read()
        
        return pdf_bytes


if __name__ == '__main__':
    # Simple CLI interface for testing
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python docx_pdf_converter.py <input.docx> [output_dir]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_directory = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        result_pdf = convert_docx_to_pdf(input_file, output_directory)
        print(f"✓ Successfully converted to: {result_pdf}")
    except Exception as e:
        print(f"✗ Error: {e}")
        sys.exit(1)
