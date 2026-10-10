<#
.SYNOPSIS
    将 WPS Office 的用户级(HKCU) COM 注册同步到机器级(HKLM)，供管理员（高完整性）进程使用。
.DESCRIPTION
    背景：Windows 自 Vista 起，完整性级别高于 Medium（即"以管理员身份运行"）的进程，
    其 COM 运行时忽略 HKCU\Software\Classes 下的用户级注册，只读取 HKLM 机器级注册。
    WPS 默认按用户注册，因此在管理员权限运行的 PowerShell / 智能体工具中
    New-Object -ComObject Kwps.Application 会报 80040154（没有注册类）。

    本脚本把 KWPS/KET/KWPP.Application 三个 ProgID 及其 CLSID、LocalServer32
    从 HKCU 复制到 HKLM（同时覆盖 64 位与 32 位注册表视图），使管理员进程也能自动化 WPS。

    用法（管理员 PowerShell）：
        powershell -File register_wps_com_machine.ps1
        pwsh -File register_wps_com_machine.ps1

    回滚：删除 HKLM 下以下键（64 位与 32 位视图各一份）：
        HKLM\SOFTWARE\Classes\KWPS.Application     （KET/KWPP 同名键同理）
        HKLM\SOFTWARE\Classes\CLSID\{000209FF-0000-4b30-A977-D214852036FF} 等对应 CLSID 键

    注意：WPS 升级到新版本目录后，若机器级注册指向的旧路径失效，重新运行本脚本即可
    从 HKCU 刷新为最新路径。
#>
param(
    [string[]]$ProgIds = @('KWPS.Application', 'KET.Application', 'KWPP.Application')
)

$ErrorActionPreference = 'Continue'

function Get-ExeFromCommand([string]$cmd) {
    # 正确处理「带引号 / 不带引号但路径含空格」两种注册值，取出真实 exe 路径
    if ([string]::IsNullOrWhiteSpace($cmd)) { return $null }
    $m = [regex]::Match($cmd, '"([^"]+\.exe)"')
    if ($m.Success) { return $m.Groups[1].Value }
    $m = [regex]::Match($cmd, '^(.+?\.exe)')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

$isElevated = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole('Administrators')
if (-not $isElevated) {
    Write-Host '[ERROR] 本脚本需要管理员权限（要写入 HKLM 机器级注册表）。请在管理员 PowerShell 中重新运行。'
    return
}

$cu32 = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry32)
$cu64 = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)
$lm32 = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)
$lm64 = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)

$cuViews = @(
    [pscustomobject]@{ Key = $cu64; Tag = 'HKCU64' },
    [pscustomobject]@{ Key = $cu32; Tag = 'HKCU32' }
)
$lmViews = @(
    [pscustomobject]@{ Key = $lm64; Tag = 'HKLM64' },
    [pscustomobject]@{ Key = $lm32; Tag = 'HKLM32' }
)

Write-Host '===== 从 HKCU 读取 WPS 用户级注册并同步到 HKLM ====='
foreach ($prog in $ProgIds) {
    Write-Host ('---- ' + $prog + ' ----')

    # 1) 读取 ProgID（名称 + CLSID）
    $progName = $null; $clsid = $null; $srcTag = $null
    foreach ($v in $cuViews) {
        $k = $v.Key.OpenSubKey(('Software\Classes\' + $prog))
        if ($k) {
            $progName = [string]$k.GetValue('')
            $ck = $k.OpenSubKey('CLSID')
            if ($ck) { $clsid = [string]$ck.GetValue(''); $ck.Close() }
            $k.Close()
            if ($clsid) { $srcTag = $v.Tag; break }
        }
    }
    if (-not $clsid) { Write-Host '  [SKIP] HKCU 中未找到该 ProgID 或其 CLSID 值'; continue }

    # 2) 读取 CLSID 键内容（先查 32 位子视图——32 位 WPS 的 CLSID 通常注册在这里）
    $lsVal = $null; $clsDefault = $progName; $inVal = $null
    foreach ($v in @($cuViews[1], $cuViews[0])) {
        $g = $v.Key.OpenSubKey(('Software\Classes\CLSID\' + $clsid))
        if ($g) {
            $gd = [string]$g.GetValue('')
            if ($gd) { $clsDefault = $gd }
            $ls = $g.OpenSubKey('LocalServer32')
            if ($ls) { if (-not $lsVal) { $lsVal = [string]$ls.GetValue('') }; $ls.Close() }
            $ip = $g.OpenSubKey('InprocServer32')
            if ($ip) { if (-not $inVal) { $inVal = [string]$ip.GetValue('') }; $ip.Close() }
            $g.Close()
        }
        if ($lsVal) { break }
    }
    if (-not $lsVal) { Write-Host '  [WARN] 未找到 LocalServer32，跳过该 ProgID'; continue }

    Write-Host ('  来源视图: ' + $srcTag + '    名称: ' + $progName)
    Write-Host ('  CLSID   : ' + $clsid)
    Write-Host ('  Server  : ' + $lsVal)
    $exe = Get-ExeFromCommand $lsVal
    if ($exe) { Write-Host ('  exe     : ' + $exe + '   (exists=' + (Test-Path -LiteralPath $exe) + ')') }

    # 3) 写入 HKLM 两个视图
    foreach ($v in $lmViews) {
        $pk = $v.Key.CreateSubKey(('SOFTWARE\Classes\' + $prog))
        if ($pk) {
            $pk.SetValue('', $progName)
            $ck2 = $pk.CreateSubKey('CLSID')
            if ($ck2) { $ck2.SetValue('', $clsid); $ck2.Close() }
            $pk.Close()
        }
        $gk = $v.Key.CreateSubKey(('SOFTWARE\Classes\CLSID\' + $clsid))
        if ($gk) {
            $gk.SetValue('', $clsDefault)
            $lsk = $gk.CreateSubKey('LocalServer32')
            if ($lsk) { $lsk.SetValue('', $lsVal); $lsk.Close() }
            if ($inVal) {
                $ipk = $gk.CreateSubKey('InprocServer32')
                if ($ipk) { $ipk.SetValue('', $inVal); $ipk.Close() }
            }
            $gk.Close()
            Write-Host ('  [OK] 已写入 ' + $v.Tag)
        }
        else { Write-Host ('  [FAIL] 写入失败: ' + $v.Tag) }
    }
}

Write-Host ''
Write-Host '===== 验证（HKLM 两视图回读） ====='
foreach ($prog in $ProgIds) {
    foreach ($v in $lmViews) {
        $k = $v.Key.OpenSubKey(('SOFTWARE\Classes\' + $prog + '\CLSID'))
        if (-not $k) { Write-Host ('  [MISS] ' + $v.Tag + '  ' + $prog); continue }
        $g = [string]$k.GetValue(''); $k.Close()
        $ls = $v.Key.OpenSubKey(('SOFTWARE\Classes\CLSID\' + $g + '\LocalServer32'))
        $lv = $null
        if ($ls) { $lv = [string]$ls.GetValue(''); $ls.Close() }
        $exe = Get-ExeFromCommand $lv
        $okMark = if ($exe -and (Test-Path -LiteralPath $exe)) { 'OK' } else { 'CHECK' }
        Write-Host ('  [' + $okMark + '] ' + $v.Tag + '  ' + $prog + ' -> ' + $g + ' | ' + $exe)
    }
}
Write-Host ''
Write-Host '完成。可在管理员 PowerShell 中测试：New-Object -ComObject Kwps.Application'
