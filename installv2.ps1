# written with ChatGPT 1-9-2026
# for installation from guthub

# ==========================================
# Office Deployment Tool (ODT) - Atera
# ==========================================

$ErrorActionPreference = 'Stop'

# ---------- Variables ----------
$RepoOwner  = "rpietz-eor"
$RepoName   = "windows-install"
$Branch     = "main"

$ODTFile    = "setup.exe"
$ConfigFile = "o365eor.xml"

$BaseDir    = "$env:ProgramData\OfficeDeployment"
$LogFile    = "$BaseDir\OfficeInstall.log"

$ODTPath    = Join-Path $BaseDir $ODTFile
$ConfigPath = Join-Path $BaseDir $ConfigFile

$ODTUrl     = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/$ODTFile"
$ConfigUrl  = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/$ConfigFile"

# ---------- Logging ----------
function Write-Log {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp  $Message" | Tee-Object -FilePath $LogFile -Append
}

# ---------- Prep ----------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (-not (Test-Path $BaseDir)) {
    New-Item -Path $BaseDir -ItemType Directory -Force | Out-Null
}

Write-Log "Starting Office deployment via Atera"

# ---------- Download ODT ----------
Write-Log "Downloading Office Deployment Tool"
Invoke-WebRequest -Uri $ODTUrl -OutFile $ODTPath -UseBasicParsing

# ---------- Download Config ----------
Write-Log "Downloading configuration XML"
Invoke-WebRequest -Uri $ConfigUrl -OutFile $ConfigPath -UseBasicParsing

# ---------- Validation ----------
if (-not (Test-Path $ODTPath)) {
    Write-Log "ERROR: ODT download failed"
    exit 1
}

if (-not (Test-Path $ConfigPath)) {
    Write-Log "ERROR: Configuration XML download failed"
    exit 2
}

# ---------- Install ----------
Write-Log "Launching Office installer"
$Process = Start-Process -FilePath $ODTPath `
    -ArgumentList "/configure `"$ConfigPath`"" `
    -Wait `
    -PassThru `
    -WindowStyle Hidden

if ($Process.ExitCode -ne 0) {
    Write-Log "ERROR: Office install failed with exit code $($Process.ExitCode)"
    exit $Process.ExitCode
}

Write-Log "Office installation completed successfully"
exit 0
