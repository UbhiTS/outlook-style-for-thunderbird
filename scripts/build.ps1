[CmdletBinding()]
param(
  [string]$Version = "1.0.21"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $projectRoot "src"
$companionDir = Join-Path $projectRoot "companion"
$distDir = Join-Path $projectRoot "dist"
$archiveBase = "outlook-style-for-thunderbird-$Version"
$zipPath = Join-Path $distDir "$archiveBase.zip"
$xpiPath = Join-Path $distDir "$archiveBase.xpi"
$companionBase = "outlook-style-companion-$Version"
$companionZipPath = Join-Path $distDir "$companionBase.zip"
$companionXpiPath = Join-Path $distDir "$companionBase.xpi"

if (-not (Test-Path -LiteralPath (Join-Path $sourceDir "manifest.json"))) {
  throw "Source manifest not found: $sourceDir\manifest.json"
}

$manifest = Get-Content -LiteralPath (Join-Path $sourceDir "manifest.json") -Raw | ConvertFrom-Json
if ($manifest.version -ne $Version) {
  throw "Requested version '$Version' does not match manifest version '$($manifest.version)'."
}

if (Test-Path -LiteralPath $distDir) {
  Get-ChildItem -LiteralPath $distDir -File |
    Where-Object { $_.BaseName -in @($archiveBase, $companionBase) } |
    Remove-Item -Force
} else {
  New-Item -ItemType Directory -Path $distDir | Out-Null
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create entries explicitly so archive paths always use the ZIP-standard '/'
# separator. ZipFile.CreateFromDirectory() emits '\' separators on some Windows
# runtimes, which Gecko's JAR reader does not resolve as extension URLs.
$zipStream = [System.IO.File]::Open(
  $zipPath,
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
    Get-ChildItem -LiteralPath $sourceDir -Recurse -File |
      Sort-Object FullName |
      ForEach-Object {
        $relativePath = $_.FullName.Substring($sourceDir.Length).TrimStart([char[]]@('\', '/'))
        $entryName = $relativePath.Replace('\', '/')
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
          $archive,
          $_.FullName,
          $entryName,
          [System.IO.Compression.CompressionLevel]::Optimal
        ) | Out-Null
      }
  } finally {
    $archive.Dispose()
  }
} finally {
  $zipStream.Dispose()
}
Copy-Item -LiteralPath $zipPath -Destination $xpiPath

foreach ($artifact in @($xpiPath, $zipPath)) {
  $stream = [System.IO.File]::OpenRead($artifact)
  try {
    $archive = [System.IO.Compression.ZipArchive]::new(
      $stream,
      [System.IO.Compression.ZipArchiveMode]::Read,
      $false
    )
    try {
      if (-not $archive.GetEntry("manifest.json")) {
        throw "Archive is invalid: manifest.json is not at the root of $artifact"
      }
      $invalidEntry = $archive.Entries | Where-Object { $_.FullName.Contains('\') } | Select-Object -First 1
      if ($invalidEntry) {
        throw "Archive is invalid: ZIP entry uses a backslash: $($invalidEntry.FullName)"
      }
    } finally {
      $archive.Dispose()
    }
  } finally {
    $stream.Dispose()
  }
}

function New-AddonArchive {
  param(
    [Parameter(Mandatory)] [string]$InputDirectory,
    [Parameter(Mandatory)] [string]$OutputZip,
    [Parameter(Mandatory)] [string]$OutputXpi
  )

  $stream = [System.IO.File]::Open(
    $OutputZip,
    [System.IO.FileMode]::CreateNew,
    [System.IO.FileAccess]::ReadWrite,
    [System.IO.FileShare]::None
  )
  try {
    $addonArchive = [System.IO.Compression.ZipArchive]::new(
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
          [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $addonArchive,
            $_.FullName,
            $entryName,
            [System.IO.Compression.CompressionLevel]::Optimal
          ) | Out-Null
        }
    } finally {
      $addonArchive.Dispose()
    }
  } finally {
    $stream.Dispose()
  }
  Copy-Item -LiteralPath $OutputZip -Destination $OutputXpi
}

New-AddonArchive -InputDirectory $companionDir -OutputZip $companionZipPath -OutputXpi $companionXpiPath

Get-FileHash -Algorithm SHA256 -LiteralPath $xpiPath, $zipPath, $companionXpiPath, $companionZipPath |
  Select-Object Path, Hash
