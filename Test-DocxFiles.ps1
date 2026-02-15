# Test DOCX Files in Microsoft Word
# This script attempts to open DOCX files in Word to check if they're valid

param(
    [Parameter(Mandatory=$false)]
    [string]$TestPath = "scratch_out"
)

Write-Host "=" * 80
Write-Host "DOCX File Tester - Microsoft Word Validation"
Write-Host "=" * 80
Write-Host ""

# Get all .docx files in the test directory (excluding *-fixed.docx)
$files = Get-ChildItem -Path $TestPath -Filter "*.docx" -Recurse | 
         Where-Object { $_.Name -notlike "*-fixed*" } |
         Sort-Object Name

if ($files.Count -eq 0) {
    Write-Host "No DOCX files found in $TestPath"
    exit 1
}

Write-Host "Found $($files.Count) DOCX files to test`n"

# Create Word COM object
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone
} catch {
    Write-Host "ERROR: Could not create Word COM object. Is Microsoft Word installed?"
    Write-Host $_.Exception.Message
    exit 1
}

$results = @()

foreach ($file in $files) {
    $relativePath = $file.FullName.Replace((Get-Location).Path + "\", "")
    Write-Host "Testing: $relativePath ... " -NoNewline
    
    try {
        # Attempt to open the document
        $doc = $word.Documents.Open($file.FullName)
        
        # Check if document has any unresolved errors
        $hasErrors = $false
        
        # Some basic checks
        if ($doc.Paragraphs.Count -eq 0) {
            $hasErrors = $true
            $errorMsg = "Document has no paragraphs"
        }
        
        if (-not $hasErrors) {
            Write-Host "✓ OK" -ForegroundColor Green
            $results += [PSCustomObject]@{
                File = $relativePath
                Status = "OK"
                Error = ""
            }
        } else {
            Write-Host "⚠ WARNING: $errorMsg" -ForegroundColor Yellow
            $results += [PSCustomObject]@{
                File = $relativePath
                Status = "WARNING"
                Error = $errorMsg
            }
        }
        
        $doc.Close($false)
        
    } catch {
        Write-Host "✗ FAILED" -ForegroundColor Red
        $errorMsg = $_.Exception.Message
        $results += [PSCustomObject]@{
            File = $relativePath
            Status = "FAILED"
            Error = $errorMsg
        }
    }
}

# Clean up
try {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
} catch {
    # Ignore cleanup errors
}

# Display summary
Write-Host ""
Write-Host "=" * 80
Write-Host "SUMMARY"
Write-Host "=" * 80
Write-Host ""

$okCount = ($results | Where-Object { $_.Status -eq "OK" }).Count
$warnCount = ($results | Where-Object { $_.Status -eq "WARNING" }).Count  
$failCount = ($results | Where-Object { $_.Status -eq "FAILED" }).Count

Write-Host "✓ OK:      $okCount" -ForegroundColor Green
Write-Host "⚠ WARNING: $warnCount" -ForegroundColor Yellow
Write-Host "✗ FAILED:  $failCount" -ForegroundColor Red
Write-Host ""

if ($failCount -gt 0) {
    Write-Host "FAILED FILES:" -ForegroundColor Red
    $results | Where-Object { $_.Status -eq "FAILED" } | ForEach-Object {
        Write-Host "  - $($_.File)"
        Write-Host "    Error: $($_.Error)" -ForegroundColor DarkRed
    }
    Write-Host ""
}

if ($warnCount -gt 0) {
    Write-Host "WARNING FILES:" -ForegroundColor Yellow
    $results | Where-Object { $_.Status -eq "WARNING" } | ForEach-Object {
        Write-Host "  - $($_.File)"
        Write-Host "    Warning: $($_.Error)" -ForegroundColor DarkYellow
    }
    Write-Host ""
}

Write-Host "=" * 80
Write-Host "NEXT STEPS"
Write-Host "=" * 80
Write-Host ""
Write-Host "1. Manually open failed/warning files in Word to see repair prompts"
Write-Host "2. If Word repairs the file, save it as *-fixed.docx"
Write-Host "3. Extract both versions and compare XML to identify differences:"
Write-Host ""
Write-Host "   Expand-Archive -Path original.docx -Dest original/"
Write-Host "   Expand-Archive -Path original-fixed.docx -Dest original-fixed/"
Write-Host "   Compare-Object (Get-Content original/word/document.xml) ``"
Write-Host "                  (Get-Content original-fixed/word/document.xml)"
Write-Host ""
