#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Convert Word document (.docx) to HTML format
.DESCRIPTION
    Use WPS Office or Microsoft Word COM interface to perform Save As operation
    to get output consistent with WPS Save As HTML
#>

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputFile,

    [Parameter(Position=1)]
    [string]$OutputFile
)

# Set UTF8 encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Resolve input path
if (-not (Test-Path $InputFile)) {
    Write-Error "Error: File not found: $InputFile"
    exit 1
}

$InputFile = Resolve-Path $InputFile

# Determine output path
if (-not $OutputFile) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, ".html")
} else {
    $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
}

$OutputDir = [System.IO.Path]::GetDirectoryName($OutputFile)

Write-Host "Converting: $InputFile"
Write-Host "Output: $OutputFile"

# Word save format constants
$wdFormatHTML = 8
$converted = $false

# Try WPS Office
function Convert-UsingWPS {
    param($inPath, $outPath)
    $wps = $null
    $doc = $null

    try {
        Write-Host "Trying WPS Office..."
        $progs = @("Kwps.Application", "Wps.Application", "Kingsoft.WPS.Application")

        foreach ($prog in $progs) {
            try {
                $wps = New-Object -ComObject $prog -ErrorAction Stop
                Write-Host "Connected to: $prog"
                break
            } catch { continue }
        }

        if (-not $wps) { return $false }

        $wps.Visible = $false
        $doc = $wps.Documents.Open($inPath)
        $doc.SaveAs($outPath, [ref]$wdFormatHTML)

        Write-Host "[OK] WPS conversion successful"
        return $true
    }
    catch {
        Write-Host "WPS failed: $_"
        return $false
    }
    finally {
        if ($doc) {
            $doc.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        if ($wps) {
            $wps.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wps) | Out-Null
        }
    }
}

# Try Microsoft Word
function Convert-UsingWord {
    param($inPath, $outPath)
    $word = $null
    $doc = $null

    try {
        Write-Host "Trying Microsoft Word..."
        $word = New-Object -ComObject "Word.Application" -ErrorAction Stop
        Write-Host "Connected to Microsoft Word"

        $word.Visible = $false
        $word.DisplayAlerts = 0

        $doc = $word.Documents.Open($inPath)
        $doc.SaveAs([ref]$outPath, [ref]$wdFormatHTML)

        Write-Host "[OK] Word conversion successful"
        return $true
    }
    catch {
        Write-Host "Word failed: $_"
        return $false
    }
    finally {
        if ($doc) {
            $doc.Close([ref]$false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        if ($word) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }
}

# Try LibreOffice
function Convert-UsingLibreOffice {
    param($inPath, $outPath)

    try {
        Write-Host "Trying LibreOffice..."
        $paths = @(
            "C:\Program Files\LibreOffice\program\soffice.exe",
            "C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        )

        $soffice = $null
        foreach ($p in $paths) {
            if (Test-Path $p) { $soffice = $p; break }
        }

        if (-not $soffice) {
            $cmd = Get-Command "soffice" -ErrorAction SilentlyContinue
            if ($cmd) { $soffice = $cmd.Source }
        }

        if (-not $soffice) { return $false }

        $proc = Start-Process -FilePath $soffice `
            -ArgumentList @("--headless","--convert-to","html","--outdir",$OutputDir,$inPath) `
            -Wait -PassThru -WindowStyle Hidden

        if ($proc.ExitCode -eq 0) {
            $expected = [System.IO.Path]::Combine($OutputDir,
                [System.IO.Path]::GetFileNameWithoutExtension($inPath) + ".html")

            if ((Test-Path $expected) -and ($expected -ne $outPath)) {
                Move-Item -Path $expected -Destination $outPath -Force
            }
            Write-Host "[OK] LibreOffice conversion successful"
            return $true
        }
        return $false
    }
    catch {
        Write-Host "LibreOffice failed: $_"
        return $false
    }
}

# Try Pandoc
function Convert-UsingPandoc {
    param($inPath, $outPath)

    try {
        Write-Host "Trying Pandoc..."
        $cmd = Get-Command "pandoc" -ErrorAction SilentlyContinue
        if (-not $cmd) { return $false }

        $mediaDir = [System.IO.Path]::Combine($OutputDir,
            [System.IO.Path]::GetFileNameWithoutExtension($outPath) + ".files")

        & pandoc --standalone --from docx --to html `
            --extract-media="$mediaDir" --output="$outPath" "$inPath"

        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Pandoc conversion successful"
            return $true
        }
        return $false
    }
    catch {
        Write-Host "Pandoc failed: $_"
        return $false
    }
}

# Execute conversions in order
if (-not $converted) { $converted = Convert-UsingWPS $InputFile $OutputFile }
if (-not $converted) { $converted = Convert-UsingWord $InputFile $OutputFile }
if (-not $converted) { $converted = Convert-UsingPandoc $InputFile $OutputFile }
if (-not $converted) { $converted = Convert-UsingLibreOffice $InputFile $OutputFile }

# Cleanup
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

if ($converted -and (Test-Path $OutputFile)) {
    $size = (Get-Item $OutputFile).Length
    Write-Host ""
    Write-Host "========================================"
    Write-Host "[OK] Conversion successful!"
    Write-Host "[OK] Output: $OutputFile"
    Write-Host "[OK] Size: $([math]::Round($size/1KB,2)) KB"
    Write-Host "========================================"
    exit 0
} else {
    Write-Host ""
    Write-Host "========================================"
    Write-Host "[ERROR] All conversion methods failed"
    Write-Host ""
    Write-Host "Please install one of the following:"
    Write-Host "  1. WPS Office (https://www.wps.cn)"
    Write-Host "  2. Microsoft Word"
    Write-Host "  3. Pandoc (https://pandoc.org)"
    Write-Host "  4. LibreOffice (https://www.libreoffice.org)"
    Write-Host "========================================"
    exit 1
}
