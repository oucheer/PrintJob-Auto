param(
    [switch]$Launch
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$LogDir = Join-Path $Root "logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogFile = Join-Path $LogDir "setup.log"
try { Start-Transcript -Path $LogFile -Append -Force | Out-Null } catch { }

function Test-PythonRuntime {
    param(
        [Parameter(Mandatory = $true)][string]$Executable,
        [string[]]$PrefixArgs = @()
    )

    $PreviousPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $Probe = @(& $Executable @PrefixArgs -c "import sys; print(sys.executable); print('%d.%d' % sys.version_info[:2])" 2>$null)
        $ExitCode = $LASTEXITCODE
    }
    catch {
        return $null
    }
    finally {
        $ErrorActionPreference = $PreviousPreference
    }

    if ($ExitCode -ne 0 -or $Probe.Count -lt 2) { return $null }
    $VersionText = ([string]$Probe[-1]).Trim()
    $ExePath = ([string]$Probe[-2]).Trim()
    try { $Version = [version]::Parse($VersionText) } catch { return $null }

    # pywin32 wheels and the application's Windows dependencies are validated
    # on CPython 3.11-3.13. Do not silently accept a newer untested ABI.
    if ($Version -lt [version]"3.11" -or $Version -ge [version]"3.14") { return $null }
    if (-not (Test-Path -LiteralPath $ExePath)) { return $null }
    return $ExePath
}

function Find-PythonRuntime {
    $Launcher = Get-Command py -ErrorAction SilentlyContinue
    if ($Launcher) {
        foreach ($Version in @("3.13", "3.12", "3.11")) {
            $Found = Test-PythonRuntime -Executable $Launcher.Source -PrefixArgs @("-$Version")
            if ($Found) { return $Found }
        }
        $Found = Test-PythonRuntime -Executable $Launcher.Source -PrefixArgs @("-3")
        if ($Found) { return $Found }
    }

    foreach ($CommandName in @("python", "python3")) {
        $PythonCommand = Get-Command $CommandName -ErrorAction SilentlyContinue
        if ($PythonCommand) {
            $Found = Test-PythonRuntime -Executable $PythonCommand.Source
            if ($Found) { return $Found }
        }
    }

    # Detect standard per-user and machine-wide installation locations even
    # when the PATH has not been refreshed after a previous installation.
    $CandidateDirs = @()
    if ($env:LOCALAPPDATA) {
        $CandidateDirs += Join-Path $env:LOCALAPPDATA "Programs\Python"
    }
    if ($env:ProgramFiles) {
        $CandidateDirs += Join-Path $env:ProgramFiles "Python"
    }
    foreach ($Dir in $CandidateDirs) {
        if (Test-Path $Dir) {
            $CandidateFiles = Get-ChildItem -Path $Dir -Filter "python.exe" -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending
            foreach ($Candidate in $CandidateFiles) {
                $Found = Test-PythonRuntime -Executable $Candidate.FullName
                if ($Found) { return $Found }
            }
        }
    }
    return $null
}

function Install-PythonRuntime {
    Write-Host "No supported Python runtime found. Installing Python 3.13.17 for the current Windows user..." -ForegroundColor Yellow

    $Architecture = $env:PROCESSOR_ARCHITEW6432
    if (-not $Architecture) { $Architecture = $env:PROCESSOR_ARCHITECTURE }
    if ($Architecture -match "ARM64") {
        $PythonArch = "arm64"
    }
    elseif ($Architecture -match "AMD64") {
        $PythonArch = "amd64"
    }
    else {
        throw "This installer currently requires 64-bit Windows (x64/ARM64). Detected architecture: $Architecture"
    }

    $PythonVersion = "3.13.17"
    $InstallerName = "python-$PythonVersion-$PythonArch.exe"
    $InstallerPath = Join-Path $env:TEMP $InstallerName
    $InstallerUrl = "https://www.python.org/ftp/python/$PythonVersion/$InstallerName"
    $Downloaded = $false

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Host "Downloading official Python installer: $InstallerUrl" -ForegroundColor Cyan
        Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerPath -UseBasicParsing
        if (-not (Test-Path $InstallerPath) -or (Get-Item $InstallerPath).Length -lt 10MB) {
            throw "Downloaded Python installer is missing or unexpectedly small."
        }
        $Downloaded = $true
    }
    catch {
        Write-Host "Direct Python download failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }

    if ($Downloaded) {
        $InstallArgs = @(
            "/quiet", "InstallAllUsers=0", "PrependPath=0", "Include_launcher=1",
            "Include_pip=1", "Include_tcltk=1", "Include_test=0", "Shortcuts=0"
        )
        $Process = Start-Process -FilePath $InstallerPath -ArgumentList $InstallArgs -Wait -PassThru
        if ($Process.ExitCode -ne 0) {
            Write-Host "Python installer returned exit code $($Process.ExitCode)." -ForegroundColor DarkYellow
        }
    }

    $Found = Find-PythonRuntime
    if ($Found) { return $Found }

    # Winget is an alternative route on Windows systems where python.org is blocked.
    $Winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($Winget) {
        Write-Host "Trying Windows Package Manager as a fallback..." -ForegroundColor Yellow
        $PreviousPreference = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            & $Winget.Source install --id Python.Python.3.13 --exact --scope user --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
            Write-Host "winget exit code: $LASTEXITCODE"
        }
        finally { $ErrorActionPreference = $PreviousPreference }
        $Found = Find-PythonRuntime
        if ($Found) { return $Found }
    }

    throw @"
Unable to install Python automatically.
The installer needs access to python.org (or Windows Package Manager) once to obtain Python.
Please connect to an approved network and run this BAT file again.
Official download: https://www.python.org/downloads/windows/
"@
}

function Install-PythonRequirements {
    param([string]$VenvPython)

    Write-Host "[2/3] Installing Python packages listed in requirements.txt..." -ForegroundColor Cyan
    $PipProbePreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $VenvPython -m pip --version 2>$null
        $PipAvailable = ($LASTEXITCODE -eq 0)
    }
    finally { $ErrorActionPreference = $PipProbePreference }

    if (-not $PipAvailable) {
        Write-Host "Bootstrapping pip inside the virtual environment..." -ForegroundColor Yellow
        & $VenvPython -m ensurepip --upgrade
        if ($LASTEXITCODE -ne 0) { throw "Could not initialize pip in the virtual environment." }
    }

    $PipCommonArgs = @("-m", "pip", "install", "--disable-pip-version-check", "--timeout", "30", "--retries", "2")
    $IndexCandidates = @(
        @{ Name = "current pip configuration"; Url = $null },
        @{ Name = "Tsinghua TUNA PyPI mirror"; Url = "https://pypi.tuna.tsinghua.edu.cn/simple" },
        @{ Name = "USTC PyPI mirror"; Url = "https://pypi.mirrors.ustc.edu.cn/simple" },
        @{ Name = "Aliyun PyPI mirror"; Url = "https://mirrors.aliyun.com/pypi/simple/" }
    )
    $InstallSucceeded = $false
    $PreviousPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        foreach ($Index in $IndexCandidates) {
            Write-Host "Trying package source: $($Index.Name)" -ForegroundColor Yellow
            if ($Index.Url) {
                $PipArgs = $PipCommonArgs + @("--index-url", $Index.Url, "-r", "requirements.txt")
            }
            else {
                $PipArgs = $PipCommonArgs + @("-r", "requirements.txt")
            }
            & $VenvPython @PipArgs
            if ($LASTEXITCODE -eq 0) {
                $InstallSucceeded = $true
                Write-Host "Requirements installed successfully via $($Index.Name)." -ForegroundColor Green
                break
            }
            Write-Host "Package source failed (exit code $LASTEXITCODE): $($Index.Name)" -ForegroundColor DarkYellow
        }
    }
    finally { $ErrorActionPreference = $PreviousPreference }

    if (-not $InstallSucceeded) {
        throw "Python packages could not be downloaded. Check network/proxy settings. Details are in logs\setup.log."
    }
}

try {
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " PrintCase Studio - One-click setup" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan

    Write-Host "[1/3] Checking Python 3.11-3.13..." -ForegroundColor Cyan
    $PythonExe = Find-PythonRuntime
    if (-not $PythonExe) { $PythonExe = Install-PythonRuntime }
    Write-Host "Using Python: $PythonExe" -ForegroundColor Green
    & $PythonExe --version
    if ($LASTEXITCODE -ne 0) { throw "The detected Python interpreter could not be started." }

    $Vpy = Join-Path $Root ".venv\Scripts\python.exe"
    $VenvUsable = $false
    if (Test-Path $Vpy) {
        $PreviousPreference = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            & $Vpy -c "import sys; assert (3, 11) <= sys.version_info[:2] < (3, 14); print(sys.executable)" 2>$null
            $VenvUsable = ($LASTEXITCODE -eq 0)
        }
        finally { $ErrorActionPreference = $PreviousPreference }
    }

    if (-not $VenvUsable) {
        if (Test-Path (Join-Path $Root ".venv")) {
            $Backup = Join-Path $Root (".venv.invalid-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
            Write-Host "Existing .venv is not usable on this computer; moving it to $Backup" -ForegroundColor Yellow
            Move-Item -LiteralPath (Join-Path $Root ".venv") -Destination $Backup
        }
        & $PythonExe -m venv .venv
        if ($LASTEXITCODE -ne 0) { throw "Failed to create the local Python virtual environment." }
    }

    Install-PythonRequirements -VenvPython $Vpy

    Write-Host "Verifying required Python modules..." -ForegroundColor Cyan
    & $Vpy -c "import openpyxl, win32print, PyInstaller, pymupdf, PIL, pywinauto; print('openpyxl', openpyxl.__version__); print('pywin32 import OK'); print('PyInstaller import OK'); print('pymupdf OK'); print('Pillow OK'); print('pywinauto', pywinauto.__version__)"
    if ($LASTEXITCODE -ne 0) { throw "One or more required Python modules failed verification. See logs\setup.log." }

    Write-Host "[3/3] Setup verification..." -ForegroundColor Cyan
    Write-Host "Python dependencies: OK" -ForegroundColor Green
    Write-Host "Printing engine: built-in GDI + Word/WPS COM + pywinauto UI automation (no external print tool)" -ForegroundColor Green
    Write-Host "Setup log: $LogFile" -ForegroundColor Gray
    Write-Host "Setup complete." -ForegroundColor Green

    if ($Launch) {
        Write-Host "Starting PrintCase Studio..." -ForegroundColor Cyan
        & $Vpy (Join-Path $Root "run_app.py")
        if ($LASTEXITCODE -ne 0) { throw "PrintCase Studio exited with code $LASTEXITCODE. See logs\setup.log." }
    }
    else {
        Write-Host "Double-click the one-click BAT file to install and launch the application." -ForegroundColor Green
    }
}
catch {
    Write-Host "" 
    Write-Host "SETUP FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Diagnostic log: $LogFile" -ForegroundColor Yellow
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}
finally {
    try { Stop-Transcript | Out-Null } catch { }
}
