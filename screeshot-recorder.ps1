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

# ----- Ask how many pages to capture -----
$pages = Read-Host "How many pages do you want to capture?"

if (-not ($pages -as [int]) -or $pages -lt 1) {
    Write-Host "Invalid number. Aborting."
    exit
}

# ----- Parameters -----
$outputDir = "$PSScriptRoot\captures"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Remove-Item "$outputDir\*" -Force -Recurse -ErrorAction SilentlyContinue

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

# ----- Countdown: 5 → 1 -----
Write-Host ""
Write-Host "Recording is about to start !" -ForegroundColor Cyan
Write-Host "Make sure the Google Slides presentation window is active (Present mode)." -ForegroundColor Cyan
Write-Host ""
for ($i = 5; $i -ge 1; $i--) {
    Write-Host -NoNewline "`rRecording starts in $i s..." -ForegroundColor Red
    Start-Sleep -Seconds 1
}
# Clear countdown line
Write-Host -NoNewline "`r" + (" " * 40)
Write-Host -NoNewline "`r"
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

# ===============================
#   CHOOSE OUTPUT FORMAT(S)
# ===============================

Write-Host ""
Write-Host "Select output format:" -ForegroundColor Cyan
Write-Host "  1) PDF only   (img-2-pdf.py)"
Write-Host "  2) DOCX only  (img-2-docx.py)"
Write-Host "  3) Both PDF and DOCX (both .py script)"
$choice = Read-Host "Your choice [1/2/3]"

$doPdf  = $false
$doDocx = $false

switch ($choice) {
    "1" { $doPdf  = $true }
    "2" { $doDocx = $true }
    "3" { $doPdf  = $true; $doDocx = $true }
    default {
        Write-Host "Invalid choice, defaulting to PDF only."
        $doPdf = $true
    }
}

# Default file names produced by Python scripts
$defaultPdf  = Join-Path $PSScriptRoot "result.pdf"
$defaultDocx = Join-Path $PSScriptRoot "result.docx"

$generatedPdf  = $null
$generatedDocx = $null

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

if (-not $generatedPdf -and -not $generatedDocx) {
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

} else {
    # Keep default names
    if ($generatedPdf)  { $generatedPdf  = $defaultPdf }
    if ($generatedDocx) { $generatedDocx = $defaultDocx }
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
