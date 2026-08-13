[CmdletBinding()]
param(
  [string]$Version,
  [switch]$RequireClean
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Version)) {
  $manifestPath = Join-Path $projectRoot "src\manifest.json"
  $Version = [string](Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json).version
}

Push-Location $projectRoot
try {
  & (Join-Path $PSScriptRoot "validate.ps1") | Out-Host

  $nodeCommand = Get-Command node -ErrorAction Stop
  foreach ($scriptPath in @(
    "companion/scripts/background.js",
    "companion/api/OutlookThreadView/implementation.js"
  )) {
    & $nodeCommand.Source --check $scriptPath
    if ($LASTEXITCODE -ne 0) {
      throw "JavaScript syntax validation failed: $scriptPath"
    }
  }

  $gitCommand = Get-Command git -ErrorAction Stop
  & $gitCommand.Source diff --check HEAD -- .
  if ($LASTEXITCODE -ne 0) {
    throw "git diff --check found whitespace errors."
  }

  if ($RequireClean) {
    $statusLines = @(& $gitCommand.Source status --porcelain=v1 --untracked-files=all)
    if ($LASTEXITCODE -ne 0) {
      throw "Unable to inspect the Git worktree."
    }
    if ($statusLines.Count -gt 0) {
      throw "A production release must be built from a clean Git worktree."
    }
  }

  & (Join-Path $PSScriptRoot "verify-reproducible-build.ps1") -Version $Version | Out-Host
} finally {
  Pop-Location
}

Write-Output "Production release gate passed for version $Version."
