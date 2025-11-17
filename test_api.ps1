# Quick API Test Script
# Tests both the web UI endpoint and the API endpoint

Write-Host "DOCX to PDF API Test" -ForegroundColor Cyan
Write-Host "=" * 50

# Check if app is running
Write-Host "`nChecking if Flask app is running..."
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:5005/" -Method GET -UseBasicParsing -TimeoutSec 2
    Write-Host "[OK] App is running on port 5005" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] App is not running. Please start it with: python app.py" -ForegroundColor Red
    exit 1
}

# Check if test file exists
$testFile = "C:\temp\test.docx"
if (-Not (Test-Path $testFile)) {
    Write-Host "`n[ERROR] Test file not found: $testFile" -ForegroundColor Red
    Write-Host "Please create a test DOCX file at this location" -ForegroundColor Yellow
    exit 1
}

Write-Host "[OK] Test file found: $testFile" -ForegroundColor Green

# Test API endpoint
Write-Host "`nTesting API endpoint: POST /api/convert" -ForegroundColor Cyan
Write-Host "-" * 50

try {
    $outputFile = "$env:TEMP\test_output_api.pdf"
    
    # Create multipart form data
    $boundary = [System.Guid]::NewGuid().ToString()
    $fileBin = [System.IO.File]::ReadAllBytes($testFile)
    $enc = [System.Text.Encoding]::GetEncoding("iso-8859-1")
    
    $bodyLines = @(
        "--$boundary",
        "Content-Disposition: form-data; name=`"file`"; filename=`"test.docx`"",
        "Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "",
        $enc.GetString($fileBin),
        "--$boundary--"
    )
    
    $body = $bodyLines -join "`r`n"
    
    $response = Invoke-WebRequest `
        -Uri "http://127.0.0.1:5005/api/convert" `
        -Method POST `
        -ContentType "multipart/form-data; boundary=$boundary" `
        -Body $body `
        -OutFile $outputFile
    
    if (Test-Path $outputFile) {
        $fileSize = (Get-Item $outputFile).Length
        Write-Host "[SUCCESS] PDF created ($fileSize bytes)" -ForegroundColor Green
        Write-Host "[SUCCESS] Saved to: $outputFile" -ForegroundColor Green
        Write-Host "`nOpening PDF..." -ForegroundColor Cyan
        Start-Process $outputFile
    }
} catch {
    Write-Host "[FAILED] $($_.Exception.Message)" -ForegroundColor Red
}

# Test error handling
Write-Host "`n`nTesting error handling..." -ForegroundColor Cyan
Write-Host "-" * 50

Write-Host "`nTest 1: No file uploaded"
try {
    $response = Invoke-RestMethod -Uri "http://127.0.0.1:5005/api/convert" -Method POST
    Write-Host "Response: $response" -ForegroundColor Yellow
} catch {
    $errorResponse = $_.ErrorDetails.Message | ConvertFrom-Json
    Write-Host "[OK] Got expected error: $($errorResponse.error)" -ForegroundColor Green
}

Write-Host "`n" + "=" * 50
Write-Host "Tests complete!" -ForegroundColor Cyan
