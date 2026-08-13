function Get-ManifestTemplateVersion {
  return "0.0.0"
}

function Get-ProjectVersion {
  param(
    [Parameter(Mandatory)]
    [string]$ProjectRoot
  )

  $versionPath = Join-Path $ProjectRoot "VERSION"
  if (-not (Test-Path -LiteralPath $versionPath -PathType Leaf)) {
    throw "Version file not found: $versionPath"
  }

  $version = [System.IO.File]::ReadAllText($versionPath).Trim()
  $versionPattern = '\A(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\z'
  if (-not [regex]::IsMatch($version, $versionPattern)) {
    throw "VERSION must contain one semantic version such as 1.2.3."
  }
  return $version
}

function Get-PackagedManifestBytes {
  param(
    [Parameter(Mandatory)]
    [string]$ManifestPath,

    [Parameter(Mandatory)]
    [string]$Version
  )

  $templateVersion = Get-ManifestTemplateVersion
  $content = [System.IO.File]::ReadAllText($ManifestPath)
  $manifest = $content | ConvertFrom-Json
  if ([string]$manifest.version -ne $templateVersion) {
    throw "Manifest template must use version '$templateVersion': $ManifestPath"
  }

  $pattern = '(?m)(?<prefix>"version"\s*:\s*")' +
    [regex]::Escape($templateVersion) +
    '(?<suffix>")'
  $matches = [regex]::Matches($content, $pattern)
  if ($matches.Count -ne 1) {
    throw "Manifest template must contain exactly one version field: $ManifestPath"
  }

  $match = $matches[0]
  $replacement = $match.Groups["prefix"].Value + $Version + $match.Groups["suffix"].Value
  $rendered = $content.Substring(0, $match.Index) +
    $replacement +
    $content.Substring($match.Index + $match.Length)
  return ,([System.Text.UTF8Encoding]::new($false).GetBytes($rendered))
}
