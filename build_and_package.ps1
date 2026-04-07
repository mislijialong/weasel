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

function Stop-ExecutableProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName,
        [Parameter(Mandatory = $true)]
        [string]$ExePath,
        [string]$QuitArgument = ""
    )

    if (Test-Path $ExePath) {
        try {
            if ([string]::IsNullOrWhiteSpace($QuitArgument)) {
                & $ExePath | Out-Null
            } else {
                & $ExePath $QuitArgument | Out-Null
            }
        } catch {
            Write-Host "Warning: failed to request graceful stop for ${ProcessName}: $($_.Exception.Message)"
        }
    }

    $running = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    foreach ($p in $running) {
        try {
            Stop-Process -Id $p.Id -Force -ErrorAction Stop
            Write-Host "Stopped $ProcessName (PID=$($p.Id))."
        } catch {
            Write-Host "Warning: failed to stop ${ProcessName} (PID=$($p.Id)): $($_.Exception.Message)"
        }
    }
}

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $RepoRoot

$MSBuild = Find-MSBuild
$MakeNsis = "${env:ProgramFiles(x86)}\NSIS\Bin\makensis.exe"
Require-Path $MakeNsis

Write-Host "RepoRoot: $RepoRoot"
Write-Host "MSBuild:  $MSBuild"
Write-Host "NSIS:     $MakeNsis"
Write-Host ""
Write-Host "== Stop running Weasel processes =="
Stop-ExecutableProcess -ProcessName "WeaselServer" -ExePath (Join-Path $RepoRoot "output\WeaselServer.exe") -QuitArgument "/q"
Stop-ExecutableProcess -ProcessName "WeaselDeployer" -ExePath (Join-Path $RepoRoot "output\WeaselDeployer.exe")

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
New-Item -ItemType Directory -Path "output\Win32" -Force | Out-Null

Copy-Item -Force "output\weaselx64.dll" "output\weasel.dll"
Copy-Item -Force "LICENSE.txt" "output\LICENSE.txt"
Copy-Item -Force "README.md" "output\README.txt"
Copy-Item -Force "plum\rime-install.bat" "output\rime-install.bat"
Copy-Item -Force "librime\dist_lua_x64\lib\rime.dll" "output\rime.dll"
if (Test-Path "librime\dist_lua_Win32\lib\rime.dll") {
    Copy-Item -Force "librime\dist_lua_Win32\lib\rime.dll" "output\Win32\rime.dll"
} elseif (Test-Path "librime\dist_Win32\lib\rime.dll") {
    Copy-Item -Force "librime\dist_Win32\lib\rime.dll" "output\Win32\rime.dll"
}
if (-not (Test-Path "output\Win32\WinSparkle.dll") -and (Test-Path "output\WinSparkle.dll")) {
    Copy-Item -Force "output\WinSparkle.dll" "output\Win32\WinSparkle.dll"
}

# The WebView2Loader.dll is required for AI panel. We only provide the x64 version.
# To prevent NSIS errors for the Win32 version, we just copy the x64 version as a placeholder.
if (Test-Path "third_party\webview2\pkg\build\native\x64\WebView2Loader.dll") {
    Copy-Item -Force "third_party\webview2\pkg\build\native\x64\WebView2Loader.dll" "output\WebView2Loader.dll"
    Copy-Item -Force "third_party\webview2\pkg\build\native\x64\WebView2Loader.dll" "output\Win32\WebView2Loader.dll"
}

Copy-Item -Force "librime\share\opencc\*" "output\data\opencc\"
Copy-Item -Force "third_party\webview2\pkg\build\native\x64\WebView2Loader.dll" "output\WebView2Loader.dll"

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
