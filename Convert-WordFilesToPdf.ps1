<#
.SYNOPSIS
Converts all Microsoft Word files in a folder to PDF.

.DESCRIPTION
This script uses Microsoft Word installed on Windows to batch-convert Word files
to PDF. It supports .doc, .docx, .docm, .dot, .dotx, and .dotm files.

.PARAMETER InputDirectory
Folder containing the Word files to convert.

.PARAMETER OutputDirectory
Optional folder where PDFs will be saved. If omitted, PDFs are saved beside the
source Word files.

.PARAMETER Recurse
Search subfolders under InputDirectory.

.PARAMETER Overwrite
Replace existing PDF files.

.EXAMPLE
.\Convert-WordFilesToPdf.ps1 -InputDirectory ".\Input Word Files\"

.EXAMPLE
.\Convert-WordFilesToPdf.ps1 -InputDirectory ".\Input Word Files\" -OutputDirectory ".\Output PDF Files\" -Recurse -Overwrite
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

$inputPath = (Resolve-Path -LiteralPath $InputDirectory).Path

if ($OutputDirectory) {
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
    }

    $outputPath = (Resolve-Path -LiteralPath $OutputDirectory).Path
}

$wordExtensions = @(".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm")
$searchOption = if ($Recurse) { "AllDirectories" } else { "TopDirectoryOnly" }
$files = [System.IO.Directory]::EnumerateFiles($inputPath, "*.*", $searchOption) |
    Where-Object { $wordExtensions -contains [System.IO.Path]::GetExtension($_).ToLowerInvariant() }

function Test-IsValidLegacyDoc {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Legacy .doc files are OLE Compound Files and start with D0 CF 11 E0 A1 B1 1A E1.
    $expectedHeader = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)

    try {
        $bytes = Get-Content -LiteralPath $Path -AsByteStream -TotalCount 8
        if (-not $bytes -or $bytes.Count -lt 8) {
            return $false
        }

        for ($i = 0; $i -lt 8; $i++) {
            if ($bytes[$i] -ne $expectedHeader[$i]) {
                return $false
            }
        }

        return $true
    }
    catch {
        return $false
    }
}

if (-not $files) {
    Write-Host "No Word files found in: $inputPath"
    exit 0
}

$word = $null
$converted = 0
$skipped = 0
$failed = 0

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($file in $files) {
        $sourcePath = [System.IO.Path]::GetFullPath($file)
        $sourceDirectory = [System.IO.Path]::GetDirectoryName($sourcePath)
        $pdfName = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath) + ".pdf"
        $sourceExtension = [System.IO.Path]::GetExtension($sourcePath).ToLowerInvariant()

        if ($sourceExtension -eq ".doc" -and -not (Test-IsValidLegacyDoc -Path $sourcePath)) {
            Write-Warning "Skipping invalid legacy .doc file '$sourcePath' (file content does not match .doc format)."
            $failed++
            continue
        }

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

        $document = $null

        try {
            Write-Host "Converting: $sourcePath"
            $document = $word.Documents.Open($sourcePath, $false, $true)
            $document.SaveAs([string]$pdfPath, [int]17)
            $converted++
        }
        catch {
            Write-Warning "Failed to convert '$sourcePath': $($_.Exception.Message)"
            $failed++
        }
        finally {
            if ($document) {
                $document.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
            }
        }
    }
}
finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Done."
Write-Host "Converted: $converted"
Write-Host "Skipped:   $skipped"
Write-Host "Failed:    $failed"

if ($failed -gt 0) {
    exit 1
}
