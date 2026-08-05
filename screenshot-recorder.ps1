Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Load SendKey type only once
if (-not ("SendKey" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public class SendKey {
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

    public const int KEYEVENTF_KEYDOWN = 0;
    public const int KEYEVENTF_KEYUP = 2;
}
"@
}

# ----- Source and parameters -----
$outputDir = "$PSScriptRoot\captures"
$inputPdf = $null
$sourceMode = "capture"
$delayMs = 1200 # delay between slides (ms)

# ----- Functions -----

function Take-Screenshot($path) {
    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bmp.Dispose()
}

function Press-RightArrow {
    $VK_RIGHT = 0x27
    [SendKey]::keybd_event($VK_RIGHT, 0, [SendKey]::KEYEVENTF_KEYDOWN, 0)
    Start-Sleep -Milliseconds 10
    [SendKey]::keybd_event($VK_RIGHT, 0, [SendKey]::KEYEVENTF_KEYUP, 0)
}

Write-Host ""
Write-Host "Choose the source:" -ForegroundColor Cyan
Write-Host "  1) Capture a presentation now"
Write-Host "  2) Reuse images already present in the captures folder"
Write-Host "  3) Reuse an existing PDF (existing text retained where possible)"
$sourceChoice = Read-Host "Source [1/2/3]"

switch ($sourceChoice) {
    "2" { $sourceMode = "captures" }
    "3" { $sourceMode = "pdf" }
    default { $sourceMode = "capture" }
}

if ($sourceMode -eq "capture") {
    $pages = Read-Host "How many pages do you want to capture?"
    if (-not ($pages -as [int]) -or $pages -lt 1) {
        Write-Host "Invalid number. Aborting."
        exit
    }
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Remove-Item "$outputDir\*" -Force -Recurse -ErrorAction SilentlyContinue

# ----- Countdown: 5 to 1 -----
Write-Host ""
Write-Host "Recording is about to start !" -ForegroundColor Cyan
Write-Host "Make sure the Google Slides presentation window is active (Present mode)." -ForegroundColor Cyan
Write-Host ""
for ($i = 5; $i -ge 1; $i--) {
    Write-Host -NoNewline "`rRecording starts in $i s..." -ForegroundColor Red
    Start-Sleep -Seconds 1
}
# Clear countdown line
Write-Host -NoNewline ("`r" + (" " * 40) + "`r")
# Final message
[console]::beep(1000, 500)  
Write-Host "Recording live !" -ForegroundColor Red

# ----- Capture loop -----
for ($i = 1; $i -le $pages; $i++) {

    $filename = Join-Path $outputDir ("Page_$i.png")
    Write-Host "Capture $i / $pages -> $filename"

    Take-Screenshot $filename
    Press-RightArrow

    Start-Sleep -Milliseconds $delayMs
}

Write-Host "Recording done ! $pages capture(s) saved in: $outputDir" -ForegroundColor Red
} elseif ($sourceMode -eq "captures") {
    if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) {
        Write-Host "Captures folder not found: $outputDir" -ForegroundColor Red
        exit
    }
    $existingCaptures = @(
        Get-ChildItem -LiteralPath $outputDir -File |
            Where-Object { $_.Name -match '^(?i)Page_\d+.*\.(png|jpe?g|bmp|webp)$' }
    )
    if ($existingCaptures.Count -eq 0) {
        Write-Host "No Page_N images found in: $outputDir" -ForegroundColor Red
        exit
    }
    Write-Host "Reusing $($existingCaptures.Count) existing capture(s) from: $outputDir" -ForegroundColor Green
} else {
    $rawPdfPath = Read-Host "Path to the existing PDF"
    $rawPdfPath = $rawPdfPath.Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($rawPdfPath) -or -not (Test-Path -LiteralPath $rawPdfPath -PathType Leaf)) {
        Write-Host "PDF not found: $rawPdfPath" -ForegroundColor Red
        exit
    }
    $inputPdf = (Resolve-Path -LiteralPath $rawPdfPath).Path
    if ([System.IO.Path]::GetExtension($inputPdf).ToLowerInvariant() -ne ".pdf") {
        Write-Host "The selected file is not a PDF: $inputPdf" -ForegroundColor Red
        exit
    }
    Write-Host "Existing PDF selected: $inputPdf" -ForegroundColor Green
}

# ===============================
#   CHOOSE OUTPUT FORMAT(S)
# ===============================

if ($sourceMode -eq "pdf") {
    Write-Host "Existing searchable text is reused; Tesseract runs only where needed." -ForegroundColor Cyan
    $choice = "4"
} else {
    Write-Host ""
    Write-Host "Select output format:" -ForegroundColor Cyan
    Write-Host "  1) PDF only   (img-2-pdf.py)"
    Write-Host "  2) DOCX only  (img-2-docx.py)"
    Write-Host "  3) Both PDF and DOCX (both .py script)"
    Write-Host "  4) Searchable PDF + auto table of contents (OCR)"
    $choice = Read-Host "Your choice [1/2/3/4]"
}

$doPdf  = $false
$doDocx = $false
$doOcr  = $false
$darkMode = $false

switch ($choice) {
    "1" { $doPdf  = $true }
    "2" { $doDocx = $true }
    "3" { $doPdf  = $true; $doDocx = $true }
    "4" { $doOcr  = $true }
    default {
        Write-Host "Invalid choice, defaulting to PDF only."
        $doPdf = $true
    }
}

if ($doOcr) {
    Write-Host ""
    Write-Host "Optional display mode" -ForegroundColor Cyan
    Write-Host "  1) Keep original slide colors (default)"
    Write-Host "  2) Smart dark mode for black-on-white slides"
    Write-Host "Dark mode preserves colored content and skips slides that are already dark or photographic."
    $displayChoice = Read-Host "Display [1/2]"
    if ($displayChoice -eq "2") {
        $darkMode = $true
    }
}

# Default file names produced by Python scripts
$defaultPdf  = Join-Path $PSScriptRoot "result.pdf"
$defaultDocx = Join-Path $PSScriptRoot "result.docx"
$defaultOcr  = Join-Path $PSScriptRoot "result-searchable.pdf"
if ($inputPdf -and [System.IO.Path]::GetFullPath($inputPdf) -eq [System.IO.Path]::GetFullPath($defaultOcr)) {
    $defaultOcr = Join-Path $PSScriptRoot "result-searchable-new.pdf"
    Write-Host "The source PDF is protected; output will use: $defaultOcr" -ForegroundColor Yellow
}

$generatedPdf  = $null
$generatedDocx = $null
$generatedOcr  = $null

# ----- Run img-2-pdf.py if requested -----
if ($doPdf) {
    Write-Host "`nRunning img-2-pdf.py..."
    python "$PSScriptRoot\img-2-pdf.py"

    if (Test-Path $defaultPdf) {
        $generatedPdf = $defaultPdf
        Write-Host "PDF created: $defaultPdf"
    } else {
        Write-Host "WARNING: result.pdf not found. Something went wrong in img-2-pdf.py."
    }
}

# ----- Run img-2-docx.py if requested -----
if ($doDocx) {
    Write-Host "`nRunning img-2-docx.py..."
    python "$PSScriptRoot\img-2-docx.py"

    if (Test-Path $defaultDocx) {
        $generatedDocx = $defaultDocx
        Write-Host "DOCX created: $defaultDocx"
    } else {
        Write-Host "WARNING: result.docx not found. Something went wrong in img-2-docx.py."
    }
}

# ----- Run img-2-searchable-pdf.py if requested -----
if ($doOcr) {
    Write-Host "`nRunning img-2-searchable-pdf.py (OCR + table of contents, this can take a while)..."
    $ocrArguments = @("$PSScriptRoot\img-2-searchable-pdf.py")
    if ($inputPdf) {
        $ocrArguments += "--input-pdf"
        $ocrArguments += $inputPdf
        $ocrArguments += "--output"
        $ocrArguments += $defaultOcr
    }
    if ($darkMode) {
        $ocrArguments += "--dark-mode"
    }

    $ocrExitCode = 1
    & python @ocrArguments
    $ocrExitCode = $LASTEXITCODE

    if ($ocrExitCode -eq 0 -and (Test-Path $defaultOcr)) {
        $generatedOcr = $defaultOcr
        Write-Host "Searchable PDF created: $defaultOcr"
    } else {
        Write-Host "WARNING: searchable PDF generation failed (exit code $ocrExitCode)."
    }
}

if (-not $generatedPdf -and -not $generatedDocx -and -not $generatedOcr) {
    Write-Host "No output file was generated. Exiting."
    exit
}

# ===============================
#   RENAME OUTPUT(S)
# ===============================
Write-Host ""
$baseName = Read-Host "Name your output file (without extension, leave empty to keep 'result')"

if (-not [string]::IsNullOrWhiteSpace($baseName)) {

    # ----- Rename PDF if it exists -----
    if ($generatedPdf) {
        $newPdfName = $baseName
        if (-not $newPdfName.ToLower().EndsWith(".pdf")) {
            $newPdfName += ".pdf"
        }
        $finalPdf = Join-Path $PSScriptRoot $newPdfName

        if (Test-Path $finalPdf) {
            Remove-Item $finalPdf -Force
        }

        Rename-Item -Path $generatedPdf -NewName $newPdfName
        $generatedPdf = $finalPdf
    }

    # ----- Rename DOCX if it exists -----
    if ($generatedDocx) {
        $newDocxName = $baseName
        if (-not $newDocxName.ToLower().EndsWith(".docx")) {
            $newDocxName += ".docx"
        }
        $finalDocx = Join-Path $PSScriptRoot $newDocxName

        if (Test-Path $finalDocx) {
            Remove-Item $finalDocx -Force
        }

        Rename-Item -Path $generatedDocx -NewName $newDocxName
        $generatedDocx = $finalDocx
    }

    # ----- Rename searchable PDF if it exists -----
    if ($generatedOcr) {
        $newOcrName = $baseName
        if (-not $newOcrName.ToLower().EndsWith(".pdf")) {
            $newOcrName += ".pdf"
        }
        $finalOcr = Join-Path $PSScriptRoot $newOcrName

        if ($inputPdf -and [System.IO.Path]::GetFullPath($finalOcr) -eq [System.IO.Path]::GetFullPath($inputPdf)) {
            $safeStem = [System.IO.Path]::GetFileNameWithoutExtension($newOcrName)
            $newOcrName = "$safeStem-searchable.pdf"
            $finalOcr = Join-Path $PSScriptRoot $newOcrName
            Write-Host "The source PDF will not be overwritten. Using: $newOcrName" -ForegroundColor Yellow
        }

        if ([System.IO.Path]::GetFullPath($generatedOcr) -ne [System.IO.Path]::GetFullPath($finalOcr)) {
            if (Test-Path $finalOcr) {
                Remove-Item $finalOcr -Force
            }
            Rename-Item -Path $generatedOcr -NewName $newOcrName
        }
        $generatedOcr = $finalOcr
    }

} else {
    # Keep default names
    if ($generatedPdf)  { $generatedPdf  = $defaultPdf }
    if ($generatedDocx) { $generatedDocx = $defaultDocx }
    if ($generatedOcr)  { $generatedOcr  = $defaultOcr }
}

# ===============================
#   OPEN OUTPUT(S)
# ===============================

if ($generatedPdf -and (Test-Path $generatedPdf)) {
    Write-Host "Opening PDF: $generatedPdf"
    Start-Process $generatedPdf
}

if ($generatedDocx -and (Test-Path $generatedDocx)) {
    Write-Host "Opening DOCX: $generatedDocx"
    Start-Process $generatedDocx
}

if ($generatedOcr -and (Test-Path $generatedOcr)) {
    Write-Host "Opening searchable PDF: $generatedOcr"
    Start-Process $generatedOcr
}
