[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$Version
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$distDir = Join-Path $projectRoot "dist"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-StreamSha256 {
  param([Parameter(Mandatory)] [System.IO.Stream]$Stream)

  $sha256 = [System.Security.Cryptography.SHA256]::Create()
  try {
    return [System.BitConverter]::ToString($sha256.ComputeHash($Stream)).Replace("-", "")
  } finally {
    $sha256.Dispose()
  }
}

function Assert-ArchiveMatchesSource {
  param(
    [Parameter(Mandatory)] [string]$ArchivePath,
    [Parameter(Mandatory)] [string]$SourceDirectory
  )

  if (-not (Test-Path -LiteralPath $ArchivePath -PathType Leaf)) {
    throw "Expected build artifact does not exist: $ArchivePath"
  }

  $expectedFiles = @{}
  Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File |
    ForEach-Object {
      $relativePath = $_.FullName.Substring($SourceDirectory.Length).TrimStart([char[]]@('\', '/'))
      $expectedFiles[$relativePath.Replace('\', '/')] = $_.FullName
    }

  $fileStream = [System.IO.File]::OpenRead($ArchivePath)
  try {
    $archive = [System.IO.Compression.ZipArchive]::new(
      $fileStream,
      [System.IO.Compression.ZipArchiveMode]::Read,
      $false
    )
    try {
      $entries = @($archive.Entries | Where-Object { -not $_.FullName.EndsWith('/') })
      if (-not $archive.GetEntry("manifest.json")) {
        throw "manifest.json is not at the archive root: $ArchivePath"
      }

      $invalidEntry = $entries | Where-Object { $_.FullName.Contains('\') } | Select-Object -First 1
      if ($invalidEntry) {
        throw "ZIP entry uses a backslash in $ArchivePath`: $($invalidEntry.FullName)"
      }

      $duplicateEntry = $entries |
        Group-Object FullName |
        Where-Object Count -gt 1 |
        Select-Object -First 1
      if ($duplicateEntry) {
        throw "Duplicate ZIP entry in $ArchivePath`: $($duplicateEntry.Name)"
      }

      $actualNames = @($entries.FullName | Sort-Object)
      $expectedNames = @($expectedFiles.Keys | Sort-Object)
      if (Compare-Object $expectedNames $actualNames) {
        $difference = Compare-Object $expectedNames $actualNames | Out-String
        throw "Archive entries do not match source files in $ArchivePath`n$difference"
      }

      foreach ($entry in $entries) {
        $entryStream = $entry.Open()
        try {
          $packedHash = Get-StreamSha256 -Stream $entryStream
        } finally {
          $entryStream.Dispose()
        }
        $sourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $expectedFiles[$entry.FullName]).Hash
        if ($packedHash -ne $sourceHash) {
          throw "Packed file differs from source in $ArchivePath`: $($entry.FullName)"
        }
      }

      $manifestEntry = $archive.GetEntry("manifest.json")
      $manifestStream = $manifestEntry.Open()
      $reader = [System.IO.StreamReader]::new($manifestStream)
      try {
        $packedManifest = $reader.ReadToEnd() | ConvertFrom-Json
      } finally {
        $reader.Dispose()
        $manifestStream.Dispose()
      }
      if ($packedManifest.version -ne $Version) {
        throw "Packaged version '$($packedManifest.version)' does not match '$Version' in $ArchivePath"
      }
    } finally {
      $archive.Dispose()
    }
  } finally {
    $fileStream.Dispose()
  }
}

$packages = @(
  @{
    Name = "outlook-style-for-thunderbird-$Version"
    Source = Join-Path $projectRoot "src"
  },
  @{
    Name = "outlook-style-companion-$Version"
    Source = Join-Path $projectRoot "companion"
  }
)

foreach ($package in $packages) {
  $zipPath = Join-Path $distDir "$($package.Name).zip"
  $xpiPath = Join-Path $distDir "$($package.Name).xpi"
  Assert-ArchiveMatchesSource -ArchivePath $zipPath -SourceDirectory $package.Source
  Assert-ArchiveMatchesSource -ArchivePath $xpiPath -SourceDirectory $package.Source

  $zipHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $zipPath).Hash
  $xpiHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $xpiPath).Hash
  if ($zipHash -ne $xpiHash) {
    throw "ZIP and XPI bytes differ for $($package.Name)."
  }
}

$artifactPaths = foreach ($package in $packages) {
  Join-Path $distDir "$($package.Name).xpi"
  Join-Path $distDir "$($package.Name).zip"
}

Get-FileHash -Algorithm SHA256 -LiteralPath $artifactPaths |
  Sort-Object Path |
  Select-Object Path, Hash

Write-Output "Verified both add-on packages against source for version $Version."
