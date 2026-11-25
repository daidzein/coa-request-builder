<#
.SYNOPSIS
  Installs Python 3.12 if missing, tests the backend and the GUI, builds COARequestBuilder.exe
  (with optional app_logo.png as window picture and file icon), tests the built exe, optionally signs it.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\build.ps1
  powershell -ExecutionPolicy Bypass -File .\build.ps1 -Clean
  powershell -ExecutionPolicy Bypass -File .\build.ps1 -OneDir
  powershell -ExecutionPolicy Bypass -File .\build.ps1 -Sign
  powershell -ExecutionPolicy Bypass -File .\build.ps1 -Sign -Thumbprint <thumb>
#>
[CmdletBinding()]
param(
    [switch]$Clean,
    [switch]$OneDir,
    [switch]$Sign,
    [string]$Thumbprint,
    [string]$TimestampUrl = "http://timestamp.digicert.com",
    [switch]$TrustOnThisPC
)

$ErrorActionPreference = "Stop"
$root    = Split-Path -Parent $MyInvocation.MyCommand.Path
$appName = "COARequestBuilder"
$gui     = Join-Path $root "coa_gui.py"
$logo    = Join-Path $root "app_logo.png"
$venv    = Join-Path $root "buildenv"
$pyUrl   = "https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"   # last 3.12 release with a Windows installer

function Step($msg) { Write-Host ""; Write-Host "==> $msg" -ForegroundColor Cyan }

function Find-Python312 {
    $old = $ErrorActionPreference; $ErrorActionPreference = "Continue"
    try {
        if (Get-Command py -ErrorAction SilentlyContinue) {
            $p = & py -3.12 -c "import sys; print(sys.executable)" 2>$null
            if ($LASTEXITCODE -eq 0 -and $p) { return ($p | Select-Object -First 1).Trim() }
        }
    } catch { } finally { $ErrorActionPreference = $old }
    foreach ($c in @("$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
                     "$env:ProgramFiles\Python312\python.exe")) {
        if (Test-Path $c) { return $c }
    }
    return $null
}

function Refresh-Path {
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [Environment]::GetEnvironmentVariable("Path", "User")
}

function Run($exe, [string[]]$arguments) {
    $old = $ErrorActionPreference; $ErrorActionPreference = "Continue"
    try { & $exe @arguments; $code = $LASTEXITCODE } finally { $ErrorActionPreference = $old }
    if ($code -ne 0) { throw "Command failed (exit $code): $exe $($arguments -join ' ')" }
}

function Invoke-SelfTest($exePath, $label, [string[]]$prefix = @()) {
    $res = Join-Path $env:TEMP ("coa_selftest_{0}.txt" -f [guid]::NewGuid().ToString("N"))
    $argList = @() + $prefix + @("--selftest", "`"$res`"")
    $p = Start-Process -FilePath $exePath -ArgumentList $argList -PassThru -WindowStyle Hidden
    $null = $p.Handle                                   # keeps ExitCode readable
    if (-not $p.WaitForExit(240000)) { try { $p.Kill() } catch { }; throw "$label self-test timed out." }
    $out = if (Test-Path $res) { (Get-Content $res -Raw).Trim() } else { "(no result file written)" }
    Remove-Item $res -Force -ErrorAction SilentlyContinue
    if ($p.ExitCode -ne 0 -or $out -ne "OK") { throw "$label self-test FAILED:`n$out" }
    Write-Host "$label self-test: OK" -ForegroundColor Green
}

# ---------------------------------------------------------------- checks
foreach ($f in "coa_gui.py", "coa_engine.py", "audit.py", "audit_store.py", "test_smoke.py") {
    if (-not (Test-Path (Join-Path $root $f))) { throw "$f not found in $root" }
}
if (-not (Test-Path $logo)) {
    Write-Warning "app_logo.png not found in $root - the app will draw a stand-in logo and the exe gets the default icon."
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-Location $root

if ($Clean) {
    Step "Cleaning old build output"
    foreach ($d in "buildenv", "build", "dist") { if (Test-Path (Join-Path $root $d)) { Remove-Item (Join-Path $root $d) -Recurse -Force } }
}

# ---------------------------------------------------------------- Python 3.12
Step "Looking for Python 3.12"
$python = Find-Python312
if (-not $python) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Step "Installing Python 3.12 with winget"
        $old = $ErrorActionPreference; $ErrorActionPreference = "Continue"
        & winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
        $ErrorActionPreference = $old
        Refresh-Path
        $python = Find-Python312
    }
}
if (-not $python) {
    Step "Installing Python 3.12 from python.org"
    $installer = Join-Path $env:TEMP "python-3.12.10-amd64.exe"
    Invoke-WebRequest -Uri $pyUrl -OutFile $installer -UseBasicParsing
    $sig = Get-AuthenticodeSignature $installer
    if ($sig.Status -ne "Valid" -or $sig.SignerCertificate.Subject -notmatch "Python Software Foundation") {
        throw "Downloaded installer failed the signature check ($($sig.Status)). Stopping."
    }
    $proc = Start-Process $installer -Wait -PassThru -ArgumentList `
        "/quiet", "InstallAllUsers=0", "PrependPath=1", "Include_launcher=1", "Include_pip=1", "Include_test=0"
    if ($proc.ExitCode -ne 0) { throw "Python installer exited with code $($proc.ExitCode)" }
    Remove-Item $installer -Force -ErrorAction SilentlyContinue
    Refresh-Path
    $python = Find-Python312
}
if (-not $python) { throw "Python 3.12 could not be found or installed." }
Write-Host "Using $python  ($(& $python --version))"

# ---------------------------------------------------------------- build environment
Step "Creating the isolated build environment"
if (-not (Test-Path (Join-Path $venv "Scripts\python.exe"))) { Run $python @("-m", "venv", $venv) }
$vpy = Join-Path $venv "Scripts\python.exe"
Run $vpy @("-m", "pip", "install", "--upgrade", "pip", "--quiet")
Run $vpy @("-m", "pip", "install", "--upgrade", "pandas", "openpyxl", "tzdata", "pillow", "pyinstaller", "--quiet")
& $vpy -m pip freeze | Set-Content (Join-Path $root "build_freeze.txt")    # exact versions used

# ---------------------------------------------------------------- icon from the logo
$iconArgs = @(); $dataArgs = @()
if (Test-Path $logo) {
    Step "Preparing the logo and the exe icon"
    $ico = Join-Path $root "app_logo.ico"
    $conv = "from PIL import Image; im = Image.open(r'$logo').convert('RGBA'); s = max(im.size); " +
            "c = Image.new('RGBA', (s, s), (0, 0, 0, 0)); c.paste(im, ((s - im.width) // 2, (s - im.height) // 2)); " +
            "c.save(r'$ico', sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])"
    Run $vpy @("-c", $conv)
    $iconArgs = @("--icon", $ico)
    $dataArgs = @("--add-data", "$logo;.")
}

# ---------------------------------------------------------------- tests on the source
Step "Test 1/3: backend smoke test (test_smoke.py)"
$old = $ErrorActionPreference; $ErrorActionPreference = "Continue"
$smoke = & $vpy (Join-Path $root "test_smoke.py") 2>&1 | Out-String
$smokeCode = $LASTEXITCODE
$ErrorActionPreference = $old
if ($smokeCode -ne 0 -or $smoke -notmatch "ALL OK") { Write-Host $smoke; throw "test_smoke.py failed - fix the backend before building." }
Write-Host "Backend smoke test: ALL OK" -ForegroundColor Green

Step "Test 2/3: the GUI from source"
Invoke-SelfTest $vpy "Source" @("`"$gui`"")

# ---------------------------------------------------------------- build
Step "Building $appName"
$mode = if ($OneDir) { "--onedir" } else { "--onefile" }
$pyiArgs = @("-m", "PyInstaller", "--noconfirm", "--clean", $mode, "--windowed", "--name", $appName,
             "--collect-data", "tzdata", "--hidden-import", "PIL._tkinter_finder") + $iconArgs + $dataArgs + @($gui)
Run $vpy $pyiArgs
$exe = if ($OneDir) { Join-Path $root "dist\$appName\$appName.exe" } else { Join-Path $root "dist\$appName.exe" }
if (-not (Test-Path $exe)) { throw "Build finished but $exe was not found." }

Step "Test 3/3: the built exe"
Invoke-SelfTest $exe "Built exe"

# ---------------------------------------------------------------- sign (optional)
if ($Sign) {
    Step "Signing"
    if ($Thumbprint) {
        $cert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
                Where-Object { $_.Thumbprint -eq $Thumbprint.Replace(" ", "").ToUpper() } | Select-Object -First 1
        if (-not $cert) { throw "No certificate with that thumbprint in the certificate stores." }
    } else {
        Write-Host "No thumbprint given: creating a self-signed certificate (trusted only on PCs that import it)."
        $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=COA Request Builder Internal Tools" `
                    -CertStoreLocation Cert:\CurrentUser\My -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(2)
        $cer = Join-Path $root "$appName-signing.cer"
        Export-Certificate -Cert $cert -FilePath $cer | Out-Null
        Write-Host "Public certificate saved: $cer  (import into Trusted Root + Trusted Publishers on other PCs)"
        if ($TrustOnThisPC) {
            Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
            Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\Root | Out-Null   # Windows asks you to confirm
        }
    }
    $targets = if ($OneDir) { Get-ChildItem (Join-Path $root "dist\$appName") -Filter *.exe } else { Get-Item $exe }
    foreach ($t in $targets) {
        $r = Set-AuthenticodeSignature -FilePath $t.FullName -Certificate $cert -HashAlgorithm SHA256 `
                 -TimestampServer $TimestampUrl
        if ($r.Status -in "UnknownError", "HashMismatch") {
            Write-Warning "Signing with timestamp failed ($($r.StatusMessage)); retrying without a timestamp."
            $r = Set-AuthenticodeSignature -FilePath $t.FullName -Certificate $cert -HashAlgorithm SHA256
        }
        Write-Host ("{0}: {1}" -f $t.Name, $r.Status)
    }
}

# ---------------------------------------------------------------- report
Step "Result"
$info = Get-Item $exe
$hash = (Get-FileHash $exe -Algorithm SHA256).Hash
$sigState = (Get-AuthenticodeSignature $exe).Status
"{0}  {1}" -f $hash, $info.Name | Set-Content (Join-Path (Split-Path $exe) "$appName.sha256.txt")
Write-Host "File      : $exe"
Write-Host ("Size      : {0:N1} MB" -f ($info.Length / 1MB))
Write-Host "SHA256    : $hash"
Write-Host "Signature : $sigState   (NotSigned = unsigned build; UnknownError = self-signed, not trusted on this PC yet)"
Write-Host ""
Write-Host "Done." -ForegroundColor Green
