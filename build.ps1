# Builds dist\NousAtlas\NousAtlas.exe. Run from anywhere:
#   powershell -ExecutionPolicy Bypass -File build.ps1
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$py = Join-Path $root ".venv\Scripts\python.exe"

& $py -m pip install --quiet -r (Join-Path $root "requirements.txt")
& $py (Join-Path $root "make_icon.py")

# Version comes from the VERSION file
$version = (Get-Content (Join-Path $root "VERSION") -Raw).Trim()
$parts = @($version.Split(".") + @("0", "0", "0"))[0..3]
New-Item -ItemType Directory -Force (Join-Path $root "build") | Out-Null
$versionFile = Join-Path $root "build\version_info.txt"
(Get-Content (Join-Path $root "version_info.txt") -Raw).
    Replace("{VERSION_TUPLE}", ($parts -join ", ")).Replace("{VERSION}", $version) |
    Set-Content $versionFile -Encoding utf8

& $py -m PyInstaller --noconfirm --clean --windowed --name NousAtlas `
    --icon (Join-Path $root "icon.ico") `
    --version-file $versionFile `
    --add-data "$(Join-Path $root 'icon.ico');." `
    --add-data "$(Join-Path $root 'VERSION');." `
    --distpath (Join-Path $root "dist") --workpath (Join-Path $root "build") `
    --specpath (Join-Path $root "build") `
    --exclude-module PySide6.QtNetwork --exclude-module PySide6.QtQml `
    --exclude-module PySide6.QtQuick --exclude-module PySide6.QtPdf `
    (Join-Path $root "nousatlas.py")
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

# Qt drags in pieces this app never uses. Drop them.
$qt = Join-Path $root "dist\NousAtlas\_internal\PySide6"
$drop = @("opengl32sw.dll", "Qt6Quick*.dll", "Qt6Qml*.dll", "Qt6Pdf*.dll", "Qt6Network.dll",
          "QtNetwork.pyd", "Qt6VirtualKeyboard*.dll", "Qt6OpenGL*.dll", "translations",
          "plugins\tls", "plugins\platforminputcontexts", "plugins\networkinformation",
          "plugins\imageformats\qpdf.dll", "plugins\imageformats\qsvg.dll",
          "plugins\iconengines", "Qt6Svg.dll")
foreach ($d in $drop) { Remove-Item (Join-Path $qt $d) -Recurse -Force -ErrorAction SilentlyContinue }
$mb = (Get-ChildItem (Join-Path $root "dist\NousAtlas") -Recurse -File | Measure-Object Length -Sum).Sum / 1MB
Write-Host ("Built: {0} (version {1}, {2:N0} MB)" -f (Join-Path $root 'dist\NousAtlas\NousAtlas.exe'), $version, $mb)
