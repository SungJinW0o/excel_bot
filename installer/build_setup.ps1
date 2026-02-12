param(
    [string]$OutputName = "setup"
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$distDir = Join-Path $repoRoot "dist"
$buildDir = Join-Path $repoRoot "build\setup_exe"
$specPath = Join-Path $buildDir "$OutputName.spec"
$outputExe = Join-Path $distDir "$OutputName.exe"
$runBotPath = Join-Path $repoRoot "run_bot.py"
$runBotGuiPath = Join-Path $repoRoot "run_bot_gui.py"
$runBotGuiBatPath = Join-Path $repoRoot "run_bot_gui.bat"
$configPath = Join-Path $repoRoot "config.json"
$usersPath = Join-Path $repoRoot "users.json"
$excelBotPath = Join-Path $repoRoot "excel_bot"
$launchBatPath = Join-Path $repoRoot "installer\launch_excel_bot.bat"
$licensePath = Join-Path $repoRoot "LICENSE"
$installerEntryPath = Join-Path $repoRoot "installer\windows_setup.py"

if (Test-Path $buildDir) {
    Remove-Item -Recurse -Force $buildDir
}

if (Test-Path $outputExe) {
    Remove-Item -Force $outputExe
}

New-Item -ItemType Directory -Force -Path $buildDir | Out-Null
New-Item -ItemType Directory -Force -Path $distDir | Out-Null

Push-Location $repoRoot
try {
    python -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --console `
        --name $OutputName `
        --distpath $distDir `
        --workpath $buildDir `
        --specpath $buildDir `
        --add-data "${runBotPath};payload" `
        --add-data "${runBotGuiPath};payload" `
        --add-data "${runBotGuiBatPath};payload" `
        --add-data "${configPath};payload" `
        --add-data "${usersPath};payload" `
        --add-data "${excelBotPath};payload/excel_bot" `
        --add-data "${launchBatPath};payload" `
        --add-data "${licensePath};payload" `
        $installerEntryPath

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}

if (-not (Test-Path $outputExe)) {
    throw "Expected installer was not created: $outputExe"
}

$item = Get-Item $outputExe
Write-Output "Created installer: $($item.FullName)"
Write-Output "Size (MB): $([Math]::Round($item.Length / 1MB, 2))"
Write-Output "Spec file: $specPath"
