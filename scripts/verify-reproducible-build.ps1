[CmdletBinding()]
param(
  [string]$Version
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Version)) {
  $manifestPath = Join-Path $projectRoot "src\manifest.json"
  $Version = [string](Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json).version
}

$artifactNames = @(
  "outlook-style-for-thunderbird-$Version.xpi",
  "outlook-style-for-thunderbird-$Version.zip",
  "outlook-style-companion-$Version.xpi",
  "outlook-style-companion-$Version.zip",
  "SHA256SUMS.txt"
)

function Get-BuildHashes {
  $hashes = [System.Collections.Generic.SortedDictionary[string, string]]::new(
    [System.StringComparer]::Ordinal
  )
  foreach ($artifactName in $artifactNames) {
    $artifactPath = Join-Path (Join-Path $projectRoot "dist") $artifactName
    if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
      throw "Expected reproducibility artifact does not exist: $artifactPath"
    }
    $hashes.Add(
      $artifactName,
      (Get-FileHash -Algorithm SHA256 -LiteralPath $artifactPath).Hash
    )
  }
  return $hashes
}

& (Join-Path $PSScriptRoot "build.ps1") -Version $Version | Out-Host
$firstHashes = Get-BuildHashes

& (Join-Path $PSScriptRoot "build.ps1") -Version $Version | Out-Host
$secondHashes = Get-BuildHashes

foreach ($artifactName in $artifactNames) {
  if ($firstHashes[$artifactName] -ne $secondHashes[$artifactName]) {
    throw "Build is not reproducible for '$artifactName': $($firstHashes[$artifactName]) != $($secondHashes[$artifactName])"
  }
}

& (Join-Path $PSScriptRoot "verify-artifacts.ps1") -Version $Version | Out-Host
Write-Output "Two clean package passes produced identical bytes for version $Version."
