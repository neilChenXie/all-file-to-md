#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Convert Word document (.docx) to HTML format via WPS Office / Microsoft Word COM.
.DESCRIPTION
    Prefer WPS Office COM interface (Kwps.Application) to perform Save As operation,
    to get output consistent with WPS Save As HTML.
    If WPS COM is unavailable, fall back to Microsoft Word COM (Word.Application).

    If both COM connections fail, a diagnostic hint is printed:
    - Target file open in Office -> ask the user to close the WPS/Word window
                                    holding the target file and retry
                                    (an open target file blocks COM activation,
                                    New-Object / Documents.Open fails silently)
    - Office missing             -> suggest installing WPS or Microsoft Word
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

# Word save format constant (wdFormatHTML)
$wdFormatHTML = 8

Write-Host "Converting: $InputFile"
Write-Host "Output: $OutputFile"

$converted = $false

# ------------------------------ 1) Try WPS Office ------------------------------
Write-Host "Trying WPS Office..."
$wps = $null
$doc = $null
$wpsConnected = $false

try {
    $progs = @("Kwps.Application", "Wps.Application", "Kingsoft.WPS.Application")

    foreach ($prog in $progs) {
        try {
            $wps = New-Object -ComObject $prog -ErrorAction Stop
            Write-Host "Connected to: $prog"
            $wpsConnected = $true
            break
        } catch { continue }
    }

    if (-not $wpsConnected) {
        Write-Host "[WARN] WPS COM 连接失败（已尝试: $($progs -join ' / ')）"
    } else {
        $wps.Visible = $false
        $doc = $wps.Documents.Open($InputFile)
        $doc.SaveAs($OutputFile, [ref]$wdFormatHTML)
        $converted = $true
        Write-Host "[OK] WPS conversion successful"
    }
}
catch {
    Write-Host "WPS failed: $_"
}
finally {
    if ($doc) {
        try { $doc.Close($false) } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($wps) {
        try { $wps.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wps) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# --------------------------- 2) Try Microsoft Word ---------------------------
if (-not $converted) {
    Write-Host ""
    Write-Host "Trying Microsoft Word..."
    $word = $null
    $wordDoc = $null

    try {
        $word = New-Object -ComObject "Word.Application" -ErrorAction Stop
        Write-Host "Connected to Microsoft Word"

        $word.Visible = $false
        $word.DisplayAlerts = 0

        $wordDoc = $word.Documents.Open($InputFile)
        $wordDoc.SaveAs([ref]$OutputFile, [ref]$wdFormatHTML)
        $converted = $true
        Write-Host "[OK] Word conversion successful"
    }
    catch {
        Write-Host "Word failed: $_"
    }
    finally {
        if ($wordDoc) {
            try { $wordDoc.Close([ref]$false) } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordDoc) | Out-Null
        }
        if ($word) {
            try { $word.Quit() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# ---------------------------- Failure diagnostics ----------------------------
if (-not $converted -or -not (Test-Path $OutputFile)) {
    Write-Host ""
    Write-Host "[ERROR] WPS 与 Microsoft Word 均无法完成导出"

    # 诊断：当前进程是否为管理员（高完整性）——高完整性进程的 COM 运行时忽略用户级(HKCU)注册，
    # WPS 按用户安装/注册时其 COM 在此类进程中不可见（New-Object 报 80040154 没有注册类），
    # 这不代表 WPS 未安装或注册损坏（普通权限进程中可正常自动化）。
    $isElevated = $false
    try {
        $wp = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isElevated = $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { }
    if ($isElevated) {
        Write-Host ""
        Write-Host "检测到当前进程以管理员（高完整性）权限运行: Windows 会忽略用户级(HKCU) COM 注册，"
        Write-Host "若 WPS 为按用户安装/注册，其 COM 在此类进程中不可用（80040154）。处理办法（任选其一）:"
        Write-Host "  1) 先运行本技能 scripts\register_wps_com_machine.ps1（需管理员）把 WPS 注册同步到机器级后重试;"
        Write-Host "  2) 改用非管理员权限的 PowerShell 运行本脚本（无需注册表改动）。"
        Write-Host "注意: 不要据此判定 WPS 未安装/注册异常，先排除管理员权限因素。"
    }

    # 诊断：WPS / Word 是否正在运行（若其中打开着目标文件，COM 将无法激活或打开该文件，New-Object 会静默失败）
    $running = Get-Process -Name "wps", "wpspdf", "et", "wpp", "winword" -ErrorAction SilentlyContinue
    if ($running) {
        Write-Host ""
        Write-Host "可能原因: WPS / Word 程序打开着目标文件，COM 组件无法激活。"
        Write-Host "处理办法: 请关闭打开了目标文件的 WPS / Word 窗口（保险起见关闭所有窗口）后，重新运行本脚本重试。"
    } else {
        Write-Host ""
        Write-Host "可能原因: 未安装 WPS / Microsoft Word，或 COM 注册异常。"
        Write-Host "处理办法: 安装 WPS 或 Microsoft Word 后重试；或提示用户手动将 docx 转存为 PDF 后改用 pdf-to-md 技能；或改用 pandoc 回退脚本:"
        $pandocScript = Join-Path $PSScriptRoot "docx_to_html_pandoc.py"
        Write-Host ("  python `"{0}`" `"{1}`" `"{2}`"" -f $pandocScript, $InputFile, $OutputFile)
        Write-Host "  （需先安装依赖: pip install pypandoc-binary）"
    }
    exit 1
}

# Cleanup
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

if (Test-Path $OutputFile) {
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
    Write-Host "[ERROR] 脚本执行完成但未生成输出文件"
    exit 1
}
