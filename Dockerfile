# Use Python 3.11 slim image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install LibreOffice for high-fidelity DOCX to PDF conversion
# LibreOffice provides native Microsoft Office format support with full formatting preservation
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Expose port 80
EXPOSE 80

# Run with gunicorn on port 80
CMD ["gunicorn", "--bind", "0.0.0.0:80", "--timeout", "600", "--workers", "2", "app:app"]
