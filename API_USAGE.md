# DOCX to PDF API Endpoint

## Overview
The `/api/convert` endpoint provides programmatic access to DOCX-to-PDF conversion functionality for other applications.

## Endpoint Details

**URL:** `POST /api/convert`  
**Content-Type:** `multipart/form-data`  
**Authentication:** Optional (see below)

## Usage Examples

### Python (requests library)
```python
import requests

# Convert a DOCX file
with open('document.docx', 'rb') as f:
    files = {'file': f}
    response = requests.post(
        'https://your-app.azurewebsites.net/api/convert',
        files=files
    )

if response.status_code == 200:
    # Save the PDF
    with open('output.pdf', 'wb') as pdf_file:
        pdf_file.write(response.content)
    print("Conversion successful!")
else:
    print(f"Error: {response.json()}")
```

### Python (with file in memory - useful for ZIP workflows)
```python
import requests
import zipfile
import io

# Extract DOCX from ZIP
with zipfile.ZipFile('archive.zip', 'r') as zip_ref:
    docx_data = zip_ref.read('document.docx')

# Convert to PDF
files = {'file': ('document.docx', io.BytesIO(docx_data), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}
response = requests.post('https://your-app.azurewebsites.net/api/convert', files=files)

# Add PDF back to a new ZIP
if response.status_code == 200:
    with zipfile.ZipFile('output.zip', 'w') as zip_out:
        zip_out.writestr('document.pdf', response.content)
```

### cURL
```bash
curl -X POST \
  -F "file=@document.docx" \
  https://your-app.azurewebsites.net/api/convert \
  --output output.pdf
```

### JavaScript (Node.js with axios)
```javascript
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');

const form = new FormData();
form.append('file', fs.createReadStream('document.docx'));

axios.post('https://your-app.azurewebsites.net/api/convert', form, {
    headers: form.getHeaders(),
    responseType: 'arraybuffer'
})
.then(response => {
    fs.writeFileSync('output.pdf', response.data);
    console.log('Conversion successful!');
})
.catch(error => {
    console.error('Error:', error.response.data);
});
```

### JavaScript (Browser with fetch)
```javascript
async function convertDocxToPdf(file) {
    const formData = new FormData();
    formData.append('file', file);
    
    const response = await fetch('https://your-app.azurewebsites.net/api/convert', {
        method: 'POST',
        body: formData
    });
    
    if (response.ok) {
        const blob = await response.blob();
        // Download the PDF
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'output.pdf';
        a.click();
    } else {
        const error = await response.json();
        console.error('Error:', error);
    }
}
```

## Response Codes

| Code | Description |
|------|-------------|
| `200` | Success - PDF file returned in response body |
| `400` | Bad Request - Missing file or invalid format |
| `401` | Unauthorized - Invalid or missing API key (if enabled) |
| `500` | Server Error - Conversion failed |

## Error Response Format
```json
{
    "error": "Error message describing what went wrong"
}
```

## File Requirements
- **Format:** `.docx` files only
- **Size:** Maximum 16MB
- **Field Name:** Must be named `file` in the multipart form data

## Optional: API Key Authentication

To enable API key authentication, uncomment the authentication code in `app.py` and set an environment variable:

```python
# In app.py (already included but commented out):
api_key = request.headers.get('X-API-Key')
if api_key != os.environ.get('API_KEY'):
    return jsonify({'error': 'Unauthorized'}), 401
```

Then set `API_KEY` as an environment variable in Azure:
```bash
az webapp config appsettings set \
  --resource-group <your-resource-group> \
  --name <your-app-name> \
  --settings API_KEY="your-secure-api-key-here"
```

### Using API Key in Requests
```python
headers = {'X-API-Key': 'your-secure-api-key-here'}
response = requests.post(url, files=files, headers=headers)
```

## Rate Limiting (Recommended for Production)

Consider adding rate limiting to prevent abuse. Example using Flask-Limiter:

```python
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

limiter = Limiter(
    app=app,
    key_func=get_remote_address,
    default_limits=["100 per hour"]
)

@app.route('/api/convert', methods=['POST'])
@limiter.limit("20 per minute")
def api_convert_docx_to_pdf():
    # ... existing code ...
```

## Monitoring

Monitor API usage through Azure Application Insights or add custom logging:

```python
import logging

logging.info(f"API conversion request from {request.remote_addr}")
```

## Notes

- Conversion timeout is set to 120 seconds
- Files are automatically cleaned up after conversion
- Unique IDs prevent filename collisions in multi-user scenarios
- Works with the same LibreOffice engine as the web interface (high-fidelity conversion)
