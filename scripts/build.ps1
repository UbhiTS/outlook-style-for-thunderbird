[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "version.ps1")

$projectRoot = Split-Path -Parent $PSScriptRoot
$version = Get-ProjectVersion -ProjectRoot $projectRoot
$sourceDir = Join-Path $projectRoot "src"
$companionDir = Join-Path $projectRoot "companion"
$distDir = Join-Path $projectRoot "dist"

if (-not (Test-Path -LiteralPath (Join-Path $sourceDir "manifest.json") -PathType Leaf)) {
  throw "Theme manifest not found: $sourceDir\manifest.json"
}
if (-not (Test-Path -LiteralPath (Join-Path $companionDir "manifest.json") -PathType Leaf)) {
  throw "Companion manifest not found: $companionDir\manifest.json"
}

$packages = @(
  @{
    Name = "outlook-style-for-thunderbird-$version"
    Source = $sourceDir
  },
  @{
    Name = "outlook-style-companion-$version"
    Source = $companionDir
  }
)

if (-not (Test-Path -LiteralPath $distDir)) {
  New-Item -ItemType Directory -Path $distDir | Out-Null
}
foreach ($package in $packages) {
  Get-ChildItem -LiteralPath $distDir -File |
    Where-Object { $_.BaseName -eq $package.Name } |
    Remove-Item -Force
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-AddonArchive {
  param(
    [Parameter(Mandatory)] [string]$InputDirectory,
    [Parameter(Mandatory)] [string]$OutputZip,
    [Parameter(Mandatory)] [string]$OutputXpi,
    [Parameter(Mandatory)] [string]$PackageVersion
  )

  $stream = [System.IO.File]::Open(
    $OutputZip,
    [System.IO.FileMode]::CreateNew,
    [System.IO.FileAccess]::ReadWrite,
    [System.IO.FileShare]::None
  )
  try {
    $archive = [System.IO.Compression.ZipArchive]::new(
      $stream,
      [System.IO.Compression.ZipArchiveMode]::Create,
      $false
    )
    try {
      Get-ChildItem -LiteralPath $InputDirectory -Recurse -File |
        Sort-Object FullName |
        ForEach-Object {
          $relativePath = $_.FullName.Substring($InputDirectory.Length).TrimStart([char[]]@('\', '/'))
          $entryName = $relativePath.Replace('\', '/')
          if ($entryName -eq "manifest.json") {
            $entry = $archive.CreateEntry(
              $entryName,
              [System.IO.Compression.CompressionLevel]::Optimal
            )
            $entry.LastWriteTime = [System.DateTimeOffset]$_.LastWriteTime
            $entryStream = $entry.Open()
            try {
              $manifestBytes = Get-PackagedManifestBytes `
                -ManifestPath $_.FullName `
                -Version $PackageVersion
              $entryStream.Write($manifestBytes, 0, $manifestBytes.Length)
            } finally {
              $entryStream.Dispose()
            }
          } else {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
              $archive,
              $_.FullName,
              $entryName,
              [System.IO.Compression.CompressionLevel]::Optimal
            ) | Out-Null
          }
        }
    } finally {
      $archive.Dispose()
    }
  } finally {
    $stream.Dispose()
  }
  Copy-Item -LiteralPath $OutputZip -Destination $OutputXpi
}

foreach ($package in $packages) {
  $zipPath = Join-Path $distDir "$($package.Name).zip"
  $xpiPath = Join-Path $distDir "$($package.Name).xpi"
  New-AddonArchive `
    -InputDirectory $package.Source `
    -OutputZip $zipPath `
    -OutputXpi $xpiPath `
    -PackageVersion $version
}

$artifactPaths = foreach ($package in $packages) {
  Join-Path $distDir "$($package.Name).xpi"
  Join-Path $distDir "$($package.Name).zip"
}

Get-FileHash -Algorithm SHA256 -LiteralPath $artifactPaths |
  Sort-Object Path |
  Select-Object Path, Hash
