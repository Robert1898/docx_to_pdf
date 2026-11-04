<#
.SYNOPSIS
  Construieste aplicatia: verificari -> PyInstaller (onedir) -> semnare (optional) -> installer.

.DESCRIPTION
  Rezultate:
    dist\Document Converter\                      aplicatia portabila (ambele exe-uri)
    dist\installer\Document-Converter-Setup-X.exe installer-ul

  Semnare digitala (optional): seteaza UNA dintre variante inainte de rulare:
    $env:SIGN_CERT_THUMBPRINT = "<amprenta certificatului din Windows Certificate Store>"
    $env:SIGN_PFX = "C:\cale\cert.pfx"; $env:SIGN_PFX_PASSWORD = "<parola>"
  Fara ele, pasul de semnare este sarit.

.EXAMPLE
  .\build.ps1
  .\build.ps1 -SkipTests -NoInstaller
#>
param(
    [switch]$SkipTests,
    [switch]$NoInstaller
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $py)) {
    throw "Lipseste .venv. Ruleaza: python -m venv .venv; .venv\Scripts\pip install -r requirements-dev.txt"
}

function Invoke-Step([string]$Name, [scriptblock]$Block) {
    Write-Host "==> $Name" -ForegroundColor Cyan
    # Windows PowerShell 5.1 transforma mesajele de pe stderr in erori; PyInstaller isi scrie
    # jurnalul acolo, asa ca decidem doar dupa codul de iesire.
    $ErrorActionPreference = 'Continue'
    & $Block
    if ($LASTEXITCODE) { throw "$Name a esuat (cod $LASTEXITCODE)" }
}

function Find-Tool([string]$EnvName, [string[]]$Patterns) {
    $fromEnv = [Environment]::GetEnvironmentVariable($EnvName)
    if ($fromEnv -and (Test-Path $fromEnv)) { return $fromEnv }
    foreach ($pattern in $Patterns) {
        $hit = Get-ChildItem $pattern -ErrorAction SilentlyContinue | Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    return $null
}

function Sign-Files([string[]]$Files) {
    $thumb = $env:SIGN_CERT_THUMBPRINT
    $pfx = $env:SIGN_PFX
    if (-not $thumb -and -not $pfx) {
        Write-Host "    (fara certificat configurat: semnarea este sarita)" -ForegroundColor DarkGray
        return
    }
    $signtool = Find-Tool 'SIGNTOOL' @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe",
        "$env:ProgramFiles\Windows Kits\10\bin\*\x64\signtool.exe"
    )
    if (-not $signtool) { throw "Nu gasesc signtool.exe (Windows SDK). Seteaza `$env:SIGNTOOL." }
    $certArgs = if ($thumb) { @('/sha1', $thumb) } else { @('/f', $pfx, '/p', $env:SIGN_PFX_PASSWORD) }
    foreach ($file in $Files) {
        Invoke-Step "Semnez $(Split-Path $file -Leaf)" {
            & $signtool sign @certArgs /fd sha256 /tr http://timestamp.digicert.com /td sha256 $file
        }
    }
}

# numele si versiunea vin din convertor/__init__.py (singura sursa)
$meta = (& $py -c "import convertor as c; print('|'.join([c.__version__, c.APP_NAME, c.GUI_EXE, c.CLI_EXE, c.APP_ID]))").Trim()
$version, $appName, $guiExe, $cliExe, $appDataId = $meta.Split('|')
Write-Host "$appName v$version" -ForegroundColor Green

if (-not $SkipTests) {
    Invoke-Step 'ruff' { & $py -m ruff check . }
    Invoke-Step 'teste rapide' { & $py -m pytest -m "not office" -q }
}

Invoke-Step 'PyInstaller' { & $py -m PyInstaller --noconfirm --clean document-converter.spec }

$app = Join-Path $PSScriptRoot "dist\$appName"
Sign-Files @("$app\$guiExe", "$app\$cliExe")

if (-not $NoInstaller) {
    $iscc = Find-Tool 'ISCC' @(
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    if (-not $iscc) { throw "Nu gasesc Inno Setup 6 (ISCC.exe). Instaleaza-l sau foloseste -NoInstaller." }
    Invoke-Step 'Installer' {
        & $iscc "/DAppVersion=$version" "/DAppName=$appName" "/DAppExe=$guiExe" "/DCliExe=$cliExe" `
            "/DAppDataId=$appDataId" /Q 'installer\document-converter.iss'
    }
    $setup = Join-Path $PSScriptRoot "dist\installer\$($appName.Replace(' ', '-'))-Setup-$version.exe"
    Sign-Files @($setup)
    Write-Host "Installer: $setup" -ForegroundColor Green
}

Write-Host "Aplicatie portabila: $app" -ForegroundColor Green
