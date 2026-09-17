# V3
# Validate/update the selected tests through the invoke script, then run the assessment.
$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'Invoke-MaesterModernDashboard.ps1'
$shellCommand = Get-Command -Name 'pwsh' -ErrorAction SilentlyContinue
if ($null -eq $shellCommand) {
    $shellCommand = Get-Command -Name 'powershell.exe' -ErrorAction SilentlyContinue
}
if ($null -eq $shellCommand) {
    throw 'PowerShell was not found. Install PowerShell 7, then run this launcher again.'
}

Push-Location -LiteralPath $PSScriptRoot
try {
    & $shellCommand.Source -NoLogo -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
        -OutputRoot $PSScriptRoot -PromptModuleUpdate -UpdateTests -IncludePreview -OpenReport
    if ($LASTEXITCODE -ne 0) {
        throw "The assessment script exited with code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}
