<#
.SYNOPSIS
Converts all image files in a folder to PDF.

.DESCRIPTION
This script batch-converts images to PDF. It supports .jpg, .jpeg, .png, .bmp, .gif, and .tiff files.

IMPORTANT: This script requires ImageMagick to be installed and available in PATH.
If ImageMagick is not installed, the script will exit with an error.

ImageMagick Installation:
  - Windows (winget): winget install ImageMagick.ImageMagick
  - Windows (manual): https://imagemagick.org/script/download.php#windows
  - macOS (Homebrew): brew install imagemagick
  - Linux (apt): sudo apt-get install imagemagick

.PARAMETER InputDirectory
Folder containing the image files to convert.

.PARAMETER OutputDirectory
Optional folder where PDFs will be saved. If omitted, PDFs are saved beside the
source image files.
Defaults to .\Output PDF Files.

.PARAMETER Recurse
Search subfolders under InputDirectory.

.PARAMETER Overwrite
Replace existing PDF files.

.NOTES
Dependencies:
  - ImageMagick (required)

.EXAMPLE
.\Convert-ImagesToPdf.ps1 -InputDirectory ".\Input Images\"

.EXAMPLE
.\Convert-ImagesToPdf.ps1 -InputDirectory ".\Input Images\" -OutputDirectory ".\Output PDF Files\" -Recurse -Overwrite
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$InputDirectory,

    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory,

    [Parameter(Mandatory = $false)]
    [switch]$Recurse,

    [Parameter(Mandatory = $false)]
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

# Check if ImageMagick is installed
$magickCommand = Get-Command magick -ErrorAction SilentlyContinue
if (-not $magickCommand) {
    Write-Error "ImageMagick is not installed or not in PATH. Please install ImageMagick from https://imagemagick.org/script/download.php"
    exit 1
}

$inputPath = (Resolve-Path -LiteralPath $InputDirectory).Path

if ($OutputDirectory) {
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
    }

    $outputPath = (Resolve-Path -LiteralPath $OutputDirectory).Path
}

$imageExtensions = @(".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff")
$searchOption = if ($Recurse) { "AllDirectories" } else { "TopDirectoryOnly" }
$files = [System.IO.Directory]::EnumerateFiles($inputPath, "*.*", $searchOption) |
    Where-Object { $imageExtensions -contains [System.IO.Path]::GetExtension($_).ToLowerInvariant() }

if (-not $files) {
    Write-Host "No image files found in: $inputPath"
    exit 0
}

$converted = 0
$skipped = 0
$failed = 0

foreach ($file in $files) {
    $sourcePath = [System.IO.Path]::GetFullPath($file)
    $sourceDirectory = [System.IO.Path]::GetDirectoryName($sourcePath)
    $pdfName = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath) + ".pdf"

    if ($OutputDirectory) {
        $pdfPath = Join-Path -Path $outputPath -ChildPath $pdfName
    }
    else {
        $pdfPath = Join-Path -Path $sourceDirectory -ChildPath $pdfName
    }

    if ((Test-Path -LiteralPath $pdfPath) -and -not $Overwrite) {
        Write-Host "Skipping existing PDF: $pdfPath"
        $skipped++
        continue
    }

    try {
        Write-Host "Converting: $sourcePath"
        & magick "$sourcePath" "$pdfPath" 2>&1 | Out-Null
        
        if (-not (Test-Path -LiteralPath $pdfPath)) {
            throw "ImageMagick conversion completed but output file was not created."
        }

        $converted++
    }
    catch {
        Write-Warning "Failed to convert '$sourcePath': $($_.Exception.Message)"
        $failed++
    }
}

Write-Host ""
Write-Host "Done."
Write-Host "Converted: $converted"
Write-Host "Skipped:   $skipped"
Write-Host "Failed:    $failed"

if ($failed -gt 0) {
    exit 1
}
