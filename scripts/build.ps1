[CmdletBinding()]
param(
  [string]$Version
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $projectRoot "src"
$companionDir = Join-Path $projectRoot "companion"
$distDir = Join-Path $projectRoot "dist"
$themeManifestPath = Join-Path $sourceDir "manifest.json"
$companionManifestPath = Join-Path $companionDir "manifest.json"
$versionPattern = '^\d+\.\d+\.\d+(?:[.-][0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?$'
$archiveTimestamp = [System.DateTimeOffset]::new(
  1980,
  1,
  1,
  0,
  0,
  0,
  [System.TimeSpan]::Zero
)

foreach ($manifestPath in @($themeManifestPath, $companionManifestPath)) {
  if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Source manifest not found: $manifestPath"
  }
}

$themeManifest = Get-Content -LiteralPath $themeManifestPath -Raw | ConvertFrom-Json
$companionManifest = Get-Content -LiteralPath $companionManifestPath -Raw | ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace($Version)) {
  $Version = [string]$themeManifest.version
}
if ($Version -notmatch $versionPattern) {
  throw "Unsafe or unsupported package version '$Version'."
}
if ([string]$themeManifest.version -ne $Version) {
  throw "Requested version '$Version' does not match theme version '$($themeManifest.version)'."
}
if ([string]$companionManifest.version -ne $Version) {
  throw "Requested version '$Version' does not match companion version '$($companionManifest.version)'."
}

if (-not (Test-Path -LiteralPath $distDir -PathType Container)) {
  New-Item -ItemType Directory -Path $distDir | Out-Null
}

$stagingDir = Join-Path $distDir (".build-{0}-{1}" -f $PID, [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $stagingDir | Out-Null

Add-Type -AssemblyName System.IO.Compression

function Get-ArchiveFileMap {
  param([Parameter(Mandatory)] [string]$InputDirectory)

  $fileMap = [System.Collections.Generic.SortedDictionary[string, string]]::new(
    [System.StringComparer]::Ordinal
  )
  foreach ($file in Get-ChildItem -LiteralPath $InputDirectory -Recurse -Force -File) {
    if (($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
      throw "Symbolic links and other reparse points are not allowed in packages: $($file.FullName)"
    }

    $relativePath = $file.FullName.Substring($InputDirectory.Length).TrimStart([char[]]@('\', '/'))
    $entryName = $relativePath.Replace('\', '/')
    if ([string]::IsNullOrWhiteSpace($entryName) -or $entryName.StartsWith('/') -or $entryName.Contains('../')) {
      throw "Unsafe archive entry path derived from '$($file.FullName)': '$entryName'"
    }
    $fileMap.Add($entryName, $file.FullName)
  }

  if ($fileMap.Count -eq 0) {
    throw "Package source directory contains no files: $InputDirectory"
  }
  return $fileMap
}

function New-DeterministicAddonArchive {
  param(
    [Parameter(Mandatory)] [string]$InputDirectory,
    [Parameter(Mandatory)] [string]$OutputZip,
    [Parameter(Mandatory)] [string]$OutputXpi
  )

  $fileMap = Get-ArchiveFileMap -InputDirectory $InputDirectory
  $zipStream = [System.IO.File]::Open(
    $OutputZip,
    [System.IO.FileMode]::CreateNew,
    [System.IO.FileAccess]::ReadWrite,
    [System.IO.FileShare]::None
  )
  try {
    $archive = [System.IO.Compression.ZipArchive]::new(
      $zipStream,
      [System.IO.Compression.ZipArchiveMode]::Create,
      $false
    )
    try {
      foreach ($item in $fileMap.GetEnumerator()) {
        $entry = $archive.CreateEntry(
          $item.Key,
          [System.IO.Compression.CompressionLevel]::Optimal
        )
        $entry.LastWriteTime = $archiveTimestamp

        $sourceStream = [System.IO.File]::OpenRead($item.Value)
        try {
          $entryStream = $entry.Open()
          try {
            $sourceStream.CopyTo($entryStream)
          } finally {
            $entryStream.Dispose()
          }
        } finally {
          $sourceStream.Dispose()
        }
      }
    } finally {
      $archive.Dispose()
    }
  } finally {
    $zipStream.Dispose()
  }

  Copy-Item -LiteralPath $OutputZip -Destination $OutputXpi
}

function Write-ChecksumFile {
  param(
    [Parameter(Mandatory)] [string[]]$ArtifactPaths,
    [Parameter(Mandatory)] [string]$OutputPath
  )

  $hashesByName = [System.Collections.Generic.SortedDictionary[string, string]]::new(
    [System.StringComparer]::Ordinal
  )
  foreach ($artifactPath in $ArtifactPaths) {
    $artifactName = Split-Path -Leaf $artifactPath
    $hashesByName.Add(
      $artifactName,
      (Get-FileHash -Algorithm SHA256 -LiteralPath $artifactPath).Hash.ToLowerInvariant()
    )
  }

  $lines = foreach ($item in $hashesByName.GetEnumerator()) {
    "$($item.Value)  $($item.Key)"
  }
  $content = ([string[]]$lines -join "`n") + "`n"
  [System.IO.File]::WriteAllText(
    $OutputPath,
    $content,
    [System.Text.UTF8Encoding]::new($false)
  )
}

$packages = @(
  @{
    Name = "outlook-style-for-thunderbird-$Version"
    Source = $sourceDir
  },
  @{
    Name = "outlook-style-companion-$Version"
    Source = $companionDir
  }
)

try {
  $stagedArtifacts = [System.Collections.Generic.List[string]]::new()
  foreach ($package in $packages) {
    $stagedZip = Join-Path $stagingDir "$($package.Name).zip"
    $stagedXpi = Join-Path $stagingDir "$($package.Name).xpi"
    New-DeterministicAddonArchive `
      -InputDirectory $package.Source `
      -OutputZip $stagedZip `
      -OutputXpi $stagedXpi
    $stagedArtifacts.Add($stagedXpi)
    $stagedArtifacts.Add($stagedZip)
  }

  $stagedChecksum = Join-Path $stagingDir "SHA256SUMS.txt"
  Write-ChecksumFile -ArtifactPaths $stagedArtifacts.ToArray() -OutputPath $stagedChecksum

  # Every artifact is complete before any published file is replaced. The checksum
  # is moved last so it is the commit marker for a complete build set.
  foreach ($stagedArtifact in $stagedArtifacts) {
    Move-Item `
      -LiteralPath $stagedArtifact `
      -Destination (Join-Path $distDir (Split-Path -Leaf $stagedArtifact)) `
      -Force
  }
  Move-Item `
    -LiteralPath $stagedChecksum `
    -Destination (Join-Path $distDir "SHA256SUMS.txt") `
    -Force
} finally {
  if (Test-Path -LiteralPath $stagingDir -PathType Container) {
    Remove-Item -LiteralPath $stagingDir -Recurse -Force
  }
}

$publishedArtifacts = foreach ($package in $packages) {
  Join-Path $distDir "$($package.Name).xpi"
  Join-Path $distDir "$($package.Name).zip"
}

Get-FileHash -Algorithm SHA256 -LiteralPath $publishedArtifacts |
  Sort-Object Path |
  Select-Object Path, Hash
Write-Output "Wrote deterministic packages and SHA256SUMS.txt for version $Version."
