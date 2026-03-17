param(
    [string]$Configuration = "Release",
    [string]$PlatformToolset = "v145",
    [string]$WeaselVersion = "0.17.4",
    [string]$WeaselBuild = "0",
    [string]$ProductVersion = "0.17.4"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Find-MSBuild {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $found = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" 2>$null | Select-Object -First 1
        if ($found -and (Test-Path $found)) {
            return $found
        }
    }

    $candidates = @(
        "D:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) {
            return $path
        }
    }

    throw "MSBuild.exe not found. Install Visual Studio C++ build tools first."
}

function Require-Path([string]$Path) {
    if (-not (Test-Path $Path)) {
        throw "Required path not found: $Path"
    }
}

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $RepoRoot

# Ensure RC can resolve local compatibility headers such as include\afxres.h.
$projectInclude = Join-Path $RepoRoot "include"
$projectWtlInclude = Join-Path $RepoRoot "include\wtl"
$compatInclude = Join-Path $RepoRoot "output\compat-headers"
New-Item -ItemType Directory -Path $compatInclude -Force | Out-Null

$repoAfxres = Join-Path $projectInclude "afxres.h"
$compatAfxres = Join-Path $compatInclude "afxres.h"
if (-not (Test-Path $repoAfxres)) {
    @"
#ifndef _AFXRES_H_
#define _AFXRES_H_
#include <winres.h>
#endif
"@ | Set-Content -Path $compatAfxres -Encoding ASCII
}

$includeParts = @($compatInclude, $projectInclude, $projectWtlInclude, $env:INCLUDE) | Where-Object { $_ -and $_.Trim() -ne "" }
$env:INCLUDE = ($includeParts -join ";")

$tempCompatHeaders = @()
if (-not (Test-Path $repoAfxres)) {
    foreach ($relPath in @("WeaselTSF\afxres.h", "WeaselServer\afxres.h", "WeaselSetup\afxres.h")) {
        $target = Join-Path $RepoRoot $relPath
        if (-not (Test-Path $target)) {
            Copy-Item -Force $compatAfxres $target
            $tempCompatHeaders += $target
        }
    }
}

$MSBuild = Find-MSBuild
$MakeNsis = "${env:ProgramFiles(x86)}\NSIS\Bin\makensis.exe"
Require-Path $MakeNsis

Write-Host "RepoRoot: $RepoRoot"
Write-Host "MSBuild:  $MSBuild"
Write-Host "NSIS:     $MakeNsis"
try {
    Write-Host ""
    Write-Host "== Build x64 =="
    & $MSBuild "weasel.sln" /t:Build /p:Configuration=$Configuration /p:Platform=x64 /p:PlatformToolset=$PlatformToolset /v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed for weasel.sln (x64)."
    }

    Write-Host ""
    Write-Host "== Build WeaselSetup (Win32) =="
    & $MSBuild "weasel.sln" /t:WeaselSetup /p:Configuration=$Configuration /p:Platform=Win32 /p:PlatformToolset=$PlatformToolset /v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed for WeaselSetup (Win32)."
    }

    Write-Host ""
    Write-Host "== Prepare packaging files =="
    New-Item -ItemType Directory -Path "output\data\opencc" -Force | Out-Null

    Copy-Item -Force "output\weaselx64.dll" "output\weasel.dll"
    Copy-Item -Force "LICENSE.txt" "output\LICENSE.txt"
    Copy-Item -Force "README.md" "output\README.txt"
    Copy-Item -Force "plum\rime-install.bat" "output\rime-install.bat"
    Copy-Item -Force "librime\dist_x64\lib\rime.dll" "output\rime.dll"
    Copy-Item -Force "librime\share\opencc\*" "output\data\opencc\"

    $required = @(
        "output\install.nsi",
        "output\weasel.dll",
        "output\weaselx64.dll",
        "output\WeaselDeployer.exe",
        "output\WeaselServer.exe",
        "output\WeaselSetup.exe",
        "output\rime.dll",
        "output\WinSparkle.dll",
        "output\Win32\WeaselDeployer.exe",
        "output\Win32\WeaselServer.exe",
        "output\Win32\rime.dll",
        "output\Win32\WinSparkle.dll"
    )
    foreach ($f in $required) {
        Require-Path $f
    }

    Write-Host ""
    Write-Host "== Build installer =="
    & $MakeNsis /DWEASEL_VERSION=$WeaselVersion /DWEASEL_BUILD=$WeaselBuild /DPRODUCT_VERSION=$ProductVersion "output\install.nsi"
    if ($LASTEXITCODE -ne 0) {
        throw "NSIS packaging failed."
    }

    $installer = Join-Path $RepoRoot "output\archives\weasel-$ProductVersion-installer.exe"
    Require-Path $installer
    $item = Get-Item $installer
    $hash = (Get-FileHash $installer -Algorithm SHA256).Hash

    Write-Host ""
    Write-Host "== Done =="
    Write-Host "Installer : $($item.FullName)"
    Write-Host "Size      : $($item.Length) bytes"
    Write-Host "Timestamp : $($item.LastWriteTime)"
    Write-Host "SHA256    : $hash"
}
finally {
    foreach ($tempFile in $tempCompatHeaders) {
        if (Test-Path $tempFile) {
            Remove-Item -Force $tempFile
        }
    }
}
