[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$Version
)

$ErrorActionPreference = "Stop"

$versionPattern = '^\d+\.\d+\.\d+(?:[.-][0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?$'
if ($Version -notmatch $versionPattern) {
  throw "Unsafe or unsupported package version '$Version'."
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$distDir = Join-Path $projectRoot "dist"

Add-Type -AssemblyName System.IO.Compression

function Get-StreamSha256 {
  param([Parameter(Mandatory)] [System.IO.Stream]$Stream)

  $sha256 = [System.Security.Cryptography.SHA256]::Create()
  try {
    return [System.BitConverter]::ToString($sha256.ComputeHash($Stream)).Replace("-", "")
  } finally {
    $sha256.Dispose()
  }
}

function Get-ExpectedFileMap {
  param([Parameter(Mandatory)] [string]$SourceDirectory)

  $expectedFiles = [System.Collections.Generic.SortedDictionary[string, string]]::new(
    [System.StringComparer]::Ordinal
  )
  foreach ($file in Get-ChildItem -LiteralPath $SourceDirectory -Recurse -Force -File) {
    if (($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
      throw "Symbolic links and other reparse points are not allowed in packages: $($file.FullName)"
    }
    $relativePath = $file.FullName.Substring($SourceDirectory.Length).TrimStart([char[]]@('\', '/'))
    $expectedFiles.Add($relativePath.Replace('\', '/'), $file.FullName)
  }
  return $expectedFiles
}

function Assert-DeterministicTimestamp {
  param(
    [Parameter(Mandatory)] [System.IO.Compression.ZipArchiveEntry]$Entry,
    [Parameter(Mandatory)] [string]$ArchivePath
  )

  $timestamp = $Entry.LastWriteTime.DateTime
  if (
    $timestamp.Year -ne 1980 -or
    $timestamp.Month -ne 1 -or
    $timestamp.Day -ne 1 -or
    $timestamp.Hour -ne 0 -or
    $timestamp.Minute -ne 0 -or
    $timestamp.Second -ne 0
  ) {
    throw "Non-deterministic timestamp on '$($Entry.FullName)' in $ArchivePath`: $($Entry.LastWriteTime)"
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

  $expectedFiles = Get-ExpectedFileMap -SourceDirectory $SourceDirectory
  $fileStream = [System.IO.File]::OpenRead($ArchivePath)
  try {
    $archive = [System.IO.Compression.ZipArchive]::new(
      $fileStream,
      [System.IO.Compression.ZipArchiveMode]::Read,
      $false
    )
    try {
      $entries = @($archive.Entries)
      if ($entries.Count -ne $expectedFiles.Count) {
        throw "Archive entry count does not match source in $ArchivePath`: expected $($expectedFiles.Count), found $($entries.Count)."
      }
      if (-not $archive.GetEntry("manifest.json")) {
        throw "manifest.json is not at the archive root: $ArchivePath"
      }

      $seenNames = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::Ordinal
      )
      foreach ($entry in $entries) {
        $entryName = $entry.FullName
        if (
          [string]::IsNullOrWhiteSpace($entryName) -or
          $entryName.EndsWith('/') -or
          $entryName.StartsWith('/') -or
          $entryName.Contains('\') -or
          $entryName.Contains('../')
        ) {
          throw "Unsafe or unexpected ZIP entry in $ArchivePath`: '$entryName'"
        }
        if (-not $seenNames.Add($entryName)) {
          throw "Duplicate ZIP entry in $ArchivePath`: $entryName"
        }
        if (-not $expectedFiles.ContainsKey($entryName)) {
          throw "Unexpected ZIP entry in $ArchivePath`: $entryName"
        }

        Assert-DeterministicTimestamp -Entry $entry -ArchivePath $ArchivePath

        $sourcePath = $expectedFiles[$entryName]
        $sourceLength = (Get-Item -LiteralPath $sourcePath).Length
        if ($entry.Length -ne $sourceLength) {
          throw "Packed length differs from source in $ArchivePath`: $entryName"
        }

        $entryStream = $entry.Open()
        try {
          $packedHash = Get-StreamSha256 -Stream $entryStream
        } finally {
          $entryStream.Dispose()
        }
        $sourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $sourcePath).Hash
        if ($packedHash -ne $sourceHash) {
          throw "Packed file differs from source in $ArchivePath`: $entryName"
        }
      }

      foreach ($expectedName in $expectedFiles.Keys) {
        if (-not $seenNames.Contains($expectedName)) {
          throw "Missing ZIP entry in $ArchivePath`: $expectedName"
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
      if ([string]$packedManifest.version -ne $Version) {
        throw "Packaged version '$($packedManifest.version)' does not match '$Version' in $ArchivePath"
      }
    } finally {
      $archive.Dispose()
    }
  } finally {
    $fileStream.Dispose()
  }
}

function Get-ExpectedChecksumContent {
  param([Parameter(Mandatory)] [string[]]$ArtifactPaths)

  $hashesByName = [System.Collections.Generic.SortedDictionary[string, string]]::new(
    [System.StringComparer]::Ordinal
  )
  foreach ($artifactPath in $ArtifactPaths) {
    $hashesByName.Add(
      (Split-Path -Leaf $artifactPath),
      (Get-FileHash -Algorithm SHA256 -LiteralPath $artifactPath).Hash.ToLowerInvariant()
    )
  }

  $lines = foreach ($item in $hashesByName.GetEnumerator()) {
    "$($item.Value)  $($item.Key)"
  }
  return ([string[]]$lines -join "`n") + "`n"
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

$artifactPaths = [System.Collections.Generic.List[string]]::new()
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
  $artifactPaths.Add($xpiPath)
  $artifactPaths.Add($zipPath)
}

$checksumPath = Join-Path $distDir "SHA256SUMS.txt"
if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
  throw "Expected checksum file does not exist: $checksumPath"
}
$expectedChecksumContent = Get-ExpectedChecksumContent -ArtifactPaths $artifactPaths.ToArray()
$actualChecksumContent = [System.IO.File]::ReadAllText(
  $checksumPath,
  [System.Text.UTF8Encoding]::new($false)
).Replace("`r`n", "`n")
if ($actualChecksumContent -ne $expectedChecksumContent) {
  throw "SHA256SUMS.txt does not exactly match the four packaged artifacts."
}

Get-FileHash -Algorithm SHA256 -LiteralPath $artifactPaths.ToArray() |
  Sort-Object Path |
  Select-Object Path, Hash

Write-Output "Verified deterministic package contents and SHA-256 checksums for version $Version."
