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

$canonicalUtf8 = [System.Text.UTF8Encoding]::new($false, $true)
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

function Read-ExactBytes {
  param(
    [Parameter(Mandatory)] [System.IO.BinaryReader]$Reader,
    [Parameter(Mandatory)] [int]$Count,
    [Parameter(Mandatory)] [string]$Context
  )

  $bytes = $Reader.ReadBytes($Count)
  if ($bytes.Length -ne $Count) {
    throw "Unexpected end of archive while reading $Context."
  }
  return ,$bytes
}

function Get-CanonicalEntryName {
  param(
    [Parameter(Mandatory)] [System.IO.BinaryReader]$Reader,
    [Parameter(Mandatory)] [uint16]$Length,
    [Parameter(Mandatory)] [string]$Context
  )

  $bytes = Read-ExactBytes -Reader $Reader -Count $Length -Context $Context
  try {
    return $canonicalUtf8.GetString($bytes)
  } catch {
    throw "Invalid UTF-8 archive entry name while reading $Context."
  }
}

function Get-StoredDataCrc32 {
  param(
    [Parameter(Mandatory)] [System.IO.BinaryReader]$Reader,
    [Parameter(Mandatory)] [uint32]$Length,
    [Parameter(Mandatory)] [string]$Context
  )

  $crc = [uint32]::MaxValue
  $remaining = [uint64]$Length
  $buffer = [byte[]]::new(65536)
  while ($remaining -gt 0) {
    $requested = [int][Math]::Min([uint64]$buffer.Length, $remaining)
    $bytesRead = $Reader.Read($buffer, 0, $requested)
    if ($bytesRead -ne $requested) {
      throw "Unexpected end of archive while reading data for $Context."
    }
    for ($byteIndex = 0; $byteIndex -lt $bytesRead; $byteIndex++) {
      $tableIndex = [int](($crc -bxor $buffer[$byteIndex]) -band 255)
      $crc = [uint32](($crc -shr 8) -bxor $crc32Table[$tableIndex])
    }
    $remaining -= [uint64]$bytesRead
  }
  return [uint32]($crc -bxor [uint32]::MaxValue)
}

function Assert-CanonicalStoredArchive {
  param(
    [Parameter(Mandatory)] [string]$ArchivePath,
    [Parameter(Mandatory)] $ExpectedFiles
  )

  if ($ExpectedFiles.Count -gt [uint16]::MaxValue) {
    throw "Expected source contains too many entries for classic ZIP: $ArchivePath"
  }

  $stream = [System.IO.File]::OpenRead($ArchivePath)
  try {
    $reader = [System.IO.BinaryReader]::new($stream, $canonicalUtf8, $true)
    try {
      $localEntries = [System.Collections.Generic.List[object]]::new()
      foreach ($expectedName in $ExpectedFiles.Keys) {
        if ($stream.Position -gt [uint32]::MaxValue) {
          throw "Local entry offset exceeds classic ZIP limits in $ArchivePath."
        }
        $localOffset = [uint32]$stream.Position

        if ($reader.ReadUInt32() -ne [uint32]0x04034B50) {
          throw "Missing canonical local file header for '$expectedName' in $ArchivePath."
        }
        $versionNeeded = $reader.ReadUInt16()
        $flags = $reader.ReadUInt16()
        $method = $reader.ReadUInt16()
        $modifiedTime = $reader.ReadUInt16()
        $modifiedDate = $reader.ReadUInt16()
        $crc32 = $reader.ReadUInt32()
        $compressedLength = $reader.ReadUInt32()
        $uncompressedLength = $reader.ReadUInt32()
        $nameLength = $reader.ReadUInt16()
        $extraLength = $reader.ReadUInt16()
        $entryName = Get-CanonicalEntryName `
          -Reader $reader `
          -Length $nameLength `
          -Context "the local header for '$expectedName' in $ArchivePath"

        if (
          $versionNeeded -ne 20 -or
          $flags -ne 0x0800 -or
          $method -ne 0 -or
          $modifiedTime -ne 0 -or
          $modifiedDate -ne 0x0021 -or
          $extraLength -ne 0
        ) {
          throw "Non-canonical local header for '$expectedName' in $ArchivePath."
        }
        if ($entryName -cne $expectedName) {
          throw "Non-canonical entry order or name in $ArchivePath`: expected '$expectedName', found '$entryName'."
        }
        $sourceLength = (Get-Item -LiteralPath $ExpectedFiles[$expectedName]).Length
        if (
          $sourceLength -gt [uint32]::MaxValue -or
          $compressedLength -ne $sourceLength -or
          $uncompressedLength -ne $sourceLength
        ) {
          throw "Non-canonical stored length for '$entryName' in $ArchivePath."
        }

        $calculatedCrc32 = Get-StoredDataCrc32 `
          -Reader $reader `
          -Length $compressedLength `
          -Context "'$entryName' in $ArchivePath"
        if ($calculatedCrc32 -ne $crc32) {
          throw "CRC-32 mismatch for '$entryName' in $ArchivePath."
        }

        $localEntries.Add([pscustomobject]@{
          Name = $entryName
          Offset = $localOffset
          Crc32 = $crc32
          Length = $uncompressedLength
        })
      }

      if ($stream.Position -gt [uint32]::MaxValue) {
        throw "Central directory offset exceeds classic ZIP limits in $ArchivePath."
      }
      $centralDirectoryOffset = [uint32]$stream.Position

      foreach ($localEntry in $localEntries) {
        if ($reader.ReadUInt32() -ne [uint32]0x02014B50) {
          throw "Missing canonical central directory header for '$($localEntry.Name)' in $ArchivePath."
        }
        $versionMadeBy = $reader.ReadUInt16()
        $versionNeeded = $reader.ReadUInt16()
        $flags = $reader.ReadUInt16()
        $method = $reader.ReadUInt16()
        $modifiedTime = $reader.ReadUInt16()
        $modifiedDate = $reader.ReadUInt16()
        $crc32 = $reader.ReadUInt32()
        $compressedLength = $reader.ReadUInt32()
        $uncompressedLength = $reader.ReadUInt32()
        $nameLength = $reader.ReadUInt16()
        $extraLength = $reader.ReadUInt16()
        $commentLength = $reader.ReadUInt16()
        $startingDisk = $reader.ReadUInt16()
        $internalAttributes = $reader.ReadUInt16()
        $externalAttributes = $reader.ReadUInt32()
        $localOffset = $reader.ReadUInt32()
        $entryName = Get-CanonicalEntryName `
          -Reader $reader `
          -Length $nameLength `
          -Context "the central header for '$($localEntry.Name)' in $ArchivePath"

        if (
          $versionMadeBy -ne 0x0014 -or
          $versionNeeded -ne 20 -or
          $flags -ne 0x0800 -or
          $method -ne 0 -or
          $modifiedTime -ne 0 -or
          $modifiedDate -ne 0x0021 -or
          $extraLength -ne 0 -or
          $commentLength -ne 0 -or
          $startingDisk -ne 0 -or
          $internalAttributes -ne 0 -or
          $externalAttributes -ne 0
        ) {
          throw "Non-canonical central directory header for '$($localEntry.Name)' in $ArchivePath."
        }
        if (
          $entryName -cne $localEntry.Name -or
          $crc32 -ne $localEntry.Crc32 -or
          $compressedLength -ne $localEntry.Length -or
          $uncompressedLength -ne $localEntry.Length -or
          $localOffset -ne $localEntry.Offset
        ) {
          throw "Central directory does not match the local header for '$($localEntry.Name)' in $ArchivePath."
        }
      }

      $centralDirectoryLength = $stream.Position - $centralDirectoryOffset
      if ($centralDirectoryLength -gt [uint32]::MaxValue) {
        throw "Central directory length exceeds classic ZIP limits in $ArchivePath."
      }

      if ($reader.ReadUInt32() -ne [uint32]0x06054B50) {
        throw "Missing canonical end-of-central-directory record in $ArchivePath."
      }
      $currentDisk = $reader.ReadUInt16()
      $centralDirectoryDisk = $reader.ReadUInt16()
      $entriesOnDisk = $reader.ReadUInt16()
      $totalEntries = $reader.ReadUInt16()
      $recordedCentralLength = $reader.ReadUInt32()
      $recordedCentralOffset = $reader.ReadUInt32()
      $archiveCommentLength = $reader.ReadUInt16()
      if (
        $currentDisk -ne 0 -or
        $centralDirectoryDisk -ne 0 -or
        $entriesOnDisk -ne $localEntries.Count -or
        $totalEntries -ne $localEntries.Count -or
        $recordedCentralLength -ne $centralDirectoryLength -or
        $recordedCentralOffset -ne $centralDirectoryOffset -or
        $archiveCommentLength -ne 0 -or
        $stream.Position -ne $stream.Length
      ) {
        throw "Non-canonical end-of-central-directory record in $ArchivePath."
      }
    } finally {
      $reader.Dispose()
    }
  } catch [System.IO.EndOfStreamException] {
    throw "Truncated canonical ZIP structure in $ArchivePath."
  } finally {
    $stream.Dispose()
  }
}

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
  Assert-CanonicalStoredArchive -ArchivePath $ArchivePath -ExpectedFiles $expectedFiles
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
