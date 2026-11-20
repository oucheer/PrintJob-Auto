$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
if (-not (Test-Path ".venv\Scripts\python.exe")) { throw "Run setup_windows.ps1 first." }
$Vpy = Join-Path $Root ".venv\Scripts\python.exe"
& $Vpy -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) { throw "PyInstaller is missing. Rerun setup_windows.ps1 to install all project dependencies." }

# 注意：不加 --clean。某些环境（带删除保护/安全策略）会拦截 build 目录的批量清除，
# 导致打包在开始阶段就失败；改用独立工作目录即可保证干净构建。
$Work = "build_new"
$Out = "dist_new"
if (Test-Path $Work) { Remove-Item $Work -Recurse -Force -ErrorAction SilentlyContinue }
if (Test-Path $Out) { Remove-Item $Out -Recurse -Force -ErrorAction SilentlyContinue }

& $Vpy -m PyInstaller --noconfirm --onefile --windowed `
  --name PrintCaseStudio `
  --paths src `
  --workpath $Work `
  --distpath $Out `
  --add-data "templates\PrintCase_Template.xlsx;templates" `
  --add-data "templates\docs;templates\docs" `
  --add-data "README_CN.md;." `
  --hidden-import win32print --hidden-import pywintypes --hidden-import pythoncom `
  --hidden-import pywinauto --hidden-import comtypes --collect-submodules pywinauto `
  run_app.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

# EXE 正在运行时无法覆盖：先改名旧文件再拷贝，避免直接失败
if (Test-Path "dist\PrintCaseStudio.exe") {
  try { Move-Item "dist\PrintCaseStudio.exe" "dist\PrintCaseStudio_prev.exe" -Force }
  catch { throw "dist\PrintCaseStudio.exe 正在运行，请先关闭程序再重新打包。" }
}
Copy-Item "$Out\PrintCaseStudio.exe" "dist\PrintCaseStudio.exe" -Force
Remove-Item $Work, $Out -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Built: dist\PrintCaseStudio.exe" -ForegroundColor Green
Write-Host "Run dist\PrintCaseStudio.exe. Printing uses the built-in GDI engine plus Word/WPS COM; no external print tool is required."
