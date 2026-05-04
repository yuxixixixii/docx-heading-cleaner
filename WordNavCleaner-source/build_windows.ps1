$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

function Test-Python {
    param(
        [string]$Exe,
        [string[]]$Args
    )

    & $Exe @Args --version *> $null
    return $LASTEXITCODE -eq 0
}

$PythonExe = $null
$PythonArgs = @()

if (Get-Command py -ErrorAction SilentlyContinue) {
    if (Test-Python "py" @("-3.12")) {
        $PythonExe = "py"
        $PythonArgs = @("-3.12")
    } elseif (Test-Python "py" @("-3")) {
        $PythonExe = "py"
        $PythonArgs = @("-3")
    }
}

if ($null -eq $PythonExe -and (Get-Command python -ErrorAction SilentlyContinue)) {
    if (Test-Python "python" @()) {
        $PythonExe = "python"
        $PythonArgs = @()
    }
}

if ($null -eq $PythonExe) {
    throw @"
Python was not found.

Install Python 3.12 from:
  https://www.python.org/downloads/windows/

During installation, enable:
  Add python.exe to PATH

Then open a new PowerShell window and rerun this script.
"@
}

function Invoke-Python {
    param([string[]]$Args)

    & $PythonExe @PythonArgs @Args
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed: $PythonExe $($PythonArgs -join ' ') $($Args -join ' ')"
    }
}

Write-Host "Using Python:"
Invoke-Python @("--version")

Invoke-Python @("-m", "pip", "install", "--upgrade", "pip")
Invoke-Python @("-m", "pip", "install", "-r", "requirements.txt")

$IconArgs = @()
if (Test-Path ".\app_icon.ico") {
    $IconArgs = @("--icon", "app_icon.ico")
}

$BuildArgs = @(
    "-m", "PyInstaller",
    "--name", "WordNavCleaner",
    "--onefile",
    "--noconsole",
    "--clean"
)
$BuildArgs += $IconArgs
$BuildArgs += @("gui_app.py")
Invoke-Python $BuildArgs

$ExePath = Join-Path $ProjectRoot "dist\WordNavCleaner.exe"
if (-not (Test-Path $ExePath)) {
    throw "Build finished but EXE was not found at: $ExePath"
}

$RootCopy = Join-Path $ProjectRoot "WordNavCleaner.exe"
Copy-Item -Path $ExePath -Destination $RootCopy -Force

Write-Host ""
Write-Host "Build complete:"
Write-Host "  $ExePath"
Write-Host "A copy was also placed here:"
Write-Host "  $RootCopy"
