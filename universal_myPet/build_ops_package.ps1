param(
    [string]$OutDir = "",
    [string]$PackagePrefix = "universal_myPet_psi_release"
)

$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutDir) {
    $OutDir = Join-Path $projectDir "dist"
}

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$stageDir = Join-Path $OutDir ("{0}_{1}" -f $PackagePrefix, $timestamp)
$zipPath = $stageDir + ".zip"

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
if (Test-Path $stageDir) {
    Remove-Item -LiteralPath $stageDir -Recurse -Force
}
if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
New-Item -ItemType Directory -Path $stageDir | Out-Null

$robocopyArgs = @(
    $projectDir,
    $stageDir,
    "/E",
    "/XD", "__pycache__", "logs", "dist", ".git",
    "/XF", "*.pyc", "~$*.xlsm"
)

& robocopy @robocopyArgs | Out-Null
if ($LASTEXITCODE -ge 8) {
    throw "robocopy failed with code $LASTEXITCODE"
}

$logsDir = Join-Path $stageDir "logs"
$stateDir = Join-Path $stageDir "state"
New-Item -ItemType Directory -Force -Path $logsDir | Out-Null
New-Item -ItemType Directory -Force -Path $stateDir | Out-Null

Set-Content -Path (Join-Path $stageDir "token.md") -Value "# Insert JWT token here before launch" -Encoding UTF8
Set-Content -Path (Join-Path $stageDir "cookie.md") -Value "# Insert Cookie header value here before launch" -Encoding UTF8
Set-Content -Path (Join-Path $stateDir "checkpoints.json") -Value "{}" -Encoding UTF8
Set-Content -Path (Join-Path $stageDir "ROLLBACK_BODY.json") -Value "[]" -Encoding UTF8

Start-Sleep -Seconds 1
$archiveOk = $false
for ($attempt = 1; $attempt -le 5; $attempt++) {
    try {
        if (Test-Path $zipPath) {
            Remove-Item -LiteralPath $zipPath -Force
        }
        Compress-Archive -Path (Join-Path $stageDir "*") -DestinationPath $zipPath -Force
        $archiveOk = $true
        break
    } catch {
        if ($attempt -eq 5) {
            throw
        }
        Start-Sleep -Seconds 2
    }
}

Write-Host "Package prepared:"
Write-Host "  Stage: $stageDir"
Write-Host "  Zip:   $zipPath"
