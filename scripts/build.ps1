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

$crc32Polynomial = [Convert]::ToUInt32("EDB88320", 16)
$crc32Table = [uint32[]]::new(256)
for ($tableIndex = 0; $tableIndex -lt $crc32Table.Length; $tableIndex++) {
  $tableValue = [uint32]$tableIndex
  for ($bitIndex = 0; $bitIndex -lt 8; $bitIndex++) {
    if (($tableValue -band 1) -ne 0) {
      $tableValue = [uint32](($tableValue -shr 1) -bxor $crc32Polynomial)
    } else {
      $tableValue = [uint32]($tableValue -shr 1)
    }
  }
  $crc32Table[$tableIndex] = $tableValue
}

function Get-FileCrc32 {
  param([Parameter(Mandatory)] [string]$Path)

  $crc = [uint32]::MaxValue
  $buffer = [byte[]]::new(65536)
  $stream = [System.IO.File]::OpenRead($Path)
  try {
    while (($bytesRead = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
      for ($byteIndex = 0; $byteIndex -lt $bytesRead; $byteIndex++) {
        $tableIndex = [int](($crc -bxor $buffer[$byteIndex]) -band 255)
        $crc = [uint32](($crc -shr 8) -bxor $crc32Table[$tableIndex])
      }
    }
  } finally {
    $stream.Dispose()
  }
  return [uint32]($crc -bxor [uint32]::MaxValue)
}

function Write-StreamContents {
  param(
    [Parameter(Mandatory)] [System.IO.BinaryWriter]$Writer,
    [Parameter(Mandatory)] [string]$Path
  )

  $buffer = [byte[]]::new(65536)
  $stream = [System.IO.File]::OpenRead($Path)
  try {
    while (($bytesRead = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
      $Writer.Write($buffer, 0, $bytesRead)
    }
  } finally {
    $stream.Dispose()
  }
}

function New-DeterministicAddonArchive {
  param(
    [Parameter(Mandatory)] [string]$InputDirectory,
    [Parameter(Mandatory)] [string]$OutputZip,
    [Parameter(Mandatory)] [string]$OutputXpi
  )

  $fileMap = Get-ArchiveFileMap -InputDirectory $InputDirectory
  if ($fileMap.Count -gt [uint16]::MaxValue) {
    throw "Classic ZIP archives support at most $([uint16]::MaxValue) entries: $InputDirectory"
  }

  $utf8 = [System.Text.UTF8Encoding]::new($false, $true)
  $entries = [System.Collections.Generic.List[object]]::new()
  foreach ($item in $fileMap.GetEnumerator()) {
    $nameBytes = $utf8.GetBytes($item.Key)
    if ($nameBytes.Length -gt [uint16]::MaxValue) {
      throw "Archive entry name is too long for classic ZIP: $($item.Key)"
    }

    $fileLength = (Get-Item -LiteralPath $item.Value).Length
    if ($fileLength -gt [uint32]::MaxValue) {
      throw "Archive entry is too large for classic ZIP: $($item.Key)"
    }

    $entries.Add([pscustomobject]@{
      Name = $item.Key
      NameBytes = $nameBytes
      Path = $item.Value
      Length = [uint32]$fileLength
      Crc32 = Get-FileCrc32 -Path $item.Value
      LocalHeaderOffset = [uint32]0
    })
  }

  $zipStream = [System.IO.File]::Open(
    $OutputZip,
    [System.IO.FileMode]::CreateNew,
    [System.IO.FileAccess]::ReadWrite,
    [System.IO.FileShare]::None
  )
  try {
    $writer = [System.IO.BinaryWriter]::new(
      $zipStream,
      $utf8,
      $true
    )
    try {
      # Writing STORE entries ourselves avoids runtime-specific DEFLATE output.
      # Every header field is explicit so PowerShell 5.1 and PowerShell 7 produce
      # byte-for-byte identical archives without an external packaging tool.
      foreach ($entry in $entries) {
        if ($zipStream.Position -gt [uint32]::MaxValue) {
          throw "Archive is too large for classic ZIP: $OutputZip"
        }
        $entry.LocalHeaderOffset = [uint32]$zipStream.Position

        $writer.Write([uint32]0x04034B50) # Local file header signature.
        $writer.Write([uint16]20)         # ZIP 2.0.
        $writer.Write([uint16]0x0800)     # UTF-8 entry name.
        $writer.Write([uint16]0)          # STORE (no compression).
        $writer.Write([uint16]0)          # 00:00:00.
        $writer.Write([uint16]0x0021)     # 1980-01-01.
        $writer.Write([uint32]$entry.Crc32)
        $writer.Write([uint32]$entry.Length)
        $writer.Write([uint32]$entry.Length)
        $writer.Write([uint16]$entry.NameBytes.Length)
        $writer.Write([uint16]0)          # No extra field.
        $writer.Write([byte[]]$entry.NameBytes)
        Write-StreamContents -Writer $writer -Path $entry.Path
      }

      if ($zipStream.Position -gt [uint32]::MaxValue) {
        throw "Archive is too large for classic ZIP: $OutputZip"
      }
      $centralDirectoryOffset = [uint32]$zipStream.Position

      foreach ($entry in $entries) {
        $writer.Write([uint32]0x02014B50) # Central directory signature.
        $writer.Write([uint16]0x0014)     # Created by MS-DOS, ZIP 2.0.
        $writer.Write([uint16]20)         # ZIP 2.0.
        $writer.Write([uint16]0x0800)     # UTF-8 entry name.
        $writer.Write([uint16]0)          # STORE (no compression).
        $writer.Write([uint16]0)          # 00:00:00.
        $writer.Write([uint16]0x0021)     # 1980-01-01.
        $writer.Write([uint32]$entry.Crc32)
        $writer.Write([uint32]$entry.Length)
        $writer.Write([uint32]$entry.Length)
        $writer.Write([uint16]$entry.NameBytes.Length)
        $writer.Write([uint16]0)          # No extra field.
        $writer.Write([uint16]0)          # No file comment.
        $writer.Write([uint16]0)          # Starting disk.
        $writer.Write([uint16]0)          # Internal attributes.
        $writer.Write([uint32]0)          # External attributes.
        $writer.Write([uint32]$entry.LocalHeaderOffset)
        $writer.Write([byte[]]$entry.NameBytes)
      }

      $centralDirectoryLength = $zipStream.Position - $centralDirectoryOffset
      if ($centralDirectoryLength -gt [uint32]::MaxValue) {
        throw "Central directory is too large for classic ZIP: $OutputZip"
      }

      $writer.Write([uint32]0x06054B50) # End of central directory signature.
      $writer.Write([uint16]0)          # Current disk.
      $writer.Write([uint16]0)          # Central directory disk.
      $writer.Write([uint16]$entries.Count)
      $writer.Write([uint16]$entries.Count)
      $writer.Write([uint32]$centralDirectoryLength)
      $writer.Write([uint32]$centralDirectoryOffset)
      $writer.Write([uint16]0)          # No archive comment.
    } finally {
      $writer.Dispose()
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
