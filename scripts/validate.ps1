[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $projectRoot "src"
$companionDir = Join-Path $projectRoot "companion"
$manifestPath = Join-Path $sourceDir "manifest.json"
$stylePath = Join-Path $sourceDir "styles\outlook-fluent.css"
$companionManifestPath = Join-Path $companionDir "manifest.json"
$readmePath = Join-Path $projectRoot "README.md"
$expectedHomepage = "https://github.com/UbhiTS/outlook-style-for-thunderbird"
$expectedAuthor = "Outlook Style for Thunderbird contributors"
$expectedMinVersion = "153.0"
$expectedMaxVersion = "154.*"
$versionPattern = '^\d+\.\d+\.\d+(?:[.-][0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?$'

$errors = [System.Collections.Generic.List[string]]::new()

function Resolve-InPackageFile {
  param(
    [Parameter(Mandatory)] [string]$PackageRoot,
    [Parameter(Mandatory)] [string]$RelativePath,
    [Parameter(Mandatory)] [string]$Label
  )

  if ([string]::IsNullOrWhiteSpace($RelativePath) -or [System.IO.Path]::IsPathRooted($RelativePath)) {
    $errors.Add("$Label must be a non-empty relative package path: '$RelativePath'")
    return $null
  }

  $rootPath = [System.IO.Path]::GetFullPath($PackageRoot).TrimEnd([char[]]@('\', '/'))
  $candidatePath = [System.IO.Path]::GetFullPath((Join-Path $rootPath $RelativePath))
  $rootPrefix = $rootPath + [System.IO.Path]::DirectorySeparatorChar
  if (-not $candidatePath.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    $errors.Add("$Label escapes its package root: '$RelativePath'")
    return $null
  }
  if (-not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
    $errors.Add("$Label does not exist: '$RelativePath'")
    return $null
  }
  return $candidatePath
}

function Assert-NoForbiddenManifestProperties {
  param(
    [Parameter(Mandatory)] [object]$Manifest,
    [Parameter(Mandatory)] [string]$ManifestName,
    [Parameter(Mandatory)] [string[]]$ForbiddenProperties
  )

  $propertyNames = @($Manifest.PSObject.Properties.Name)
  foreach ($propertyName in $ForbiddenProperties) {
    if ($propertyName -in $propertyNames) {
      $errors.Add("$ManifestName must not declare '$propertyName'.")
    }
  }
}

foreach ($requiredPath in @($manifestPath, $stylePath, $companionManifestPath, $readmePath)) {
  if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
    $errors.Add("Required source file does not exist: $requiredPath")
  }
}
if ($errors.Count -gt 0) {
  $errors | ForEach-Object { Write-Error $_ }
  throw "Validation failed before manifests could be loaded."
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$companionManifest = Get-Content -LiteralPath $companionManifestPath -Raw | ConvertFrom-Json

foreach ($manifestRecord in @(
  @{ Name = "Theme"; Value = $manifest },
  @{ Name = "Companion"; Value = $companionManifest }
)) {
  $manifestValue = $manifestRecord.Value
  if ($manifestValue.manifest_version -ne 3) {
    $errors.Add("$($manifestRecord.Name) manifest_version must be 3.")
  }
  if ([string]$manifestValue.version -notmatch $versionPattern) {
    $errors.Add("$($manifestRecord.Name) version is unsafe or unsupported: '$($manifestValue.version)'.")
  }
  if ([string]$manifestValue.author -ne $expectedAuthor) {
    $errors.Add("$($manifestRecord.Name) author must be '$expectedAuthor'.")
  }
  if ([string]$manifestValue.homepage_url -ne $expectedHomepage) {
    $errors.Add("$($manifestRecord.Name) homepage_url must be '$expectedHomepage'.")
  }
  if ([string]$manifestValue.browser_specific_settings.gecko.strict_min_version -ne $expectedMinVersion) {
    $errors.Add("$($manifestRecord.Name) strict_min_version must be $expectedMinVersion.")
  }
  if ([string]$manifestValue.browser_specific_settings.gecko.strict_max_version -ne $expectedMaxVersion) {
    $errors.Add("$($manifestRecord.Name) strict_max_version must be $expectedMaxVersion for ATN Experiment validation.")
  }
  if (-not $manifestValue.browser_specific_settings.gecko.id) {
    $errors.Add("$($manifestRecord.Name) Gecko ID is required for permanent installation and safe updates.")
  }
}

if ([string]$manifest.browser_specific_settings.gecko.id -ne "{c53d41f1-6d13-4a72-b250-b089f6ba56a1}") {
  $errors.Add("The stable theme Gecko ID changed unexpectedly.")
}
if ([string]$companionManifest.browser_specific_settings.gecko.id -ne "fluent-mail-light-message-view@local") {
  $errors.Add("The stable companion Gecko ID changed unexpectedly.")
}
if ([string]$companionManifest.version -ne [string]$manifest.version) {
  $errors.Add("Theme and companion versions must match.")
}

$version = [string]$manifest.version
$escapedVersion = [System.Text.RegularExpressions.Regex]::Escape($version)
$readme = Get-Content -LiteralPath $readmePath -Raw
if ($readme -notmatch "(?m)^Current package version: \*\*$escapedVersion\*\*\.$") {
  $errors.Add("README.md must declare the current package version as $version.")
}
if ($readme -notmatch "(?m)^## Version $escapedVersion(?:\s|$)") {
  $errors.Add("README.md must contain release notes for version $version.")
}
foreach ($artifactPrefix in @("outlook-style-for-thunderbird", "outlook-style-companion")) {
  if ($readme -notmatch [System.Text.RegularExpressions.Regex]::Escape("$artifactPrefix-$version.xpi")) {
    $errors.Add("README.md installation instructions must reference $artifactPrefix-$version.xpi.")
  }
}

Assert-NoForbiddenManifestProperties `
  -Manifest $manifest `
  -ManifestName "Static theme" `
  -ForbiddenProperties @(
    "background",
    "content_scripts",
    "externally_connectable",
    "host_permissions",
    "optional_host_permissions",
    "optional_permissions",
    "permissions",
    "update_url",
    "web_accessible_resources"
  )
Assert-NoForbiddenManifestProperties `
  -Manifest $companionManifest `
  -ManifestName "Privileged companion" `
  -ForbiddenProperties @(
    "content_scripts",
    "externally_connectable",
    "host_permissions",
    "message_display_scripts",
    "optional_host_permissions",
    "optional_permissions",
    "permissions",
    "update_url",
    "web_accessible_resources"
  )

if (-not $manifest.theme.colors.frame -or -not $manifest.theme.colors.tab_background_text) {
  $errors.Add("The static theme must define frame and tab_background_text colors.")
}

$themeExperimentStyle = [string]$manifest.theme_experiment.stylesheet
if ($themeExperimentStyle) {
  Resolve-InPackageFile `
    -PackageRoot $sourceDir `
    -RelativePath $themeExperimentStyle `
    -Label "Theme experiment stylesheet" | Out-Null
} else {
  $errors.Add("theme_experiment.stylesheet is required for the detailed Outlook styling.")
}

foreach ($icon in $manifest.icons.PSObject.Properties.Value) {
  Resolve-InPackageFile `
    -PackageRoot $sourceDir `
    -RelativePath ([string]$icon) `
    -Label "Theme icon" | Out-Null
}
foreach ($icon in $companionManifest.icons.PSObject.Properties.Value) {
  Resolve-InPackageFile `
    -PackageRoot $companionDir `
    -RelativePath ([string]$icon) `
    -Label "Companion icon" | Out-Null
}

if (-not $companionManifest.background.scripts) {
  $errors.Add("The companion background registration script is required.")
}
$backgroundScripts = @($companionManifest.background.scripts)
if ($backgroundScripts.Count -ne 1 -or [string]$backgroundScripts[0] -ne "scripts/background.js") {
  $errors.Add("The companion must register only scripts/background.js as its background script.")
}
foreach ($backgroundScript in $backgroundScripts) {
  Resolve-InPackageFile `
    -PackageRoot $companionDir `
    -RelativePath ([string]$backgroundScript) `
    -Label "Companion background script" | Out-Null
}

$experimentProperties = @($companionManifest.experiment_apis.PSObject.Properties)
if (
  $experimentProperties.Count -ne 1 -or
  [string]$experimentProperties[0].Name -ne "outlookThreadView"
) {
  $errors.Add("The companion must expose only the reviewed outlookThreadView Experiment API.")
}
foreach ($experimentProperty in $experimentProperties) {
  $experiment = $experimentProperty.Value
  $experimentSchemaPath = Resolve-InPackageFile `
    -PackageRoot $companionDir `
    -RelativePath ([string]$experiment.schema) `
    -Label "Experiment schema '$($experimentProperty.Name)'"
  Resolve-InPackageFile `
    -PackageRoot $companionDir `
    -RelativePath ([string]$experiment.parent.script) `
    -Label "Experiment implementation '$($experimentProperty.Name)'" | Out-Null

  $experimentScopes = @($experiment.parent.scopes)
  if ($experimentScopes.Count -ne 1 -or $experimentScopes[0] -ne "addon_parent") {
    $errors.Add("Experiment '$($experimentProperty.Name)' must be limited to the addon_parent scope.")
  }

  if ($experimentSchemaPath) {
    $schemaDefinitions = @(Get-Content -LiteralPath $experimentSchemaPath -Raw | ConvertFrom-Json)
    if (
      $schemaDefinitions.Count -ne 1 -or
      [string]$schemaDefinitions[0].namespace -ne "outlookThreadView"
    ) {
      $errors.Add("The Experiment schema must contain only the outlookThreadView namespace.")
    } else {
      $expectedFunctions = @(
        "installCalendarChooser",
        "installEditorSurfaces",
        "installReminderDialog",
        "installTodayPane",
        "showParentThread"
      )
      $actualFunctions = @($schemaDefinitions[0].functions.name | Sort-Object)
      if (Compare-Object $expectedFunctions $actualFunctions) {
        $errors.Add("The Experiment API function surface changed and requires an explicit security review.")
      }
      foreach ($schemaFunction in @($schemaDefinitions[0].functions)) {
        if ($schemaFunction.type -ne "function" -or -not $schemaFunction.async) {
          $errors.Add("Experiment function '$($schemaFunction.name)' must remain asynchronous and function-only.")
        }
        $parameters = @($schemaFunction.parameters)
        if ($schemaFunction.name -eq "showParentThread") {
          if (
            $parameters.Count -ne 1 -or
            $parameters[0].name -ne "tabId" -or
            $parameters[0].type -ne "integer"
          ) {
            $errors.Add("showParentThread must accept only an integer tabId.")
          }
        } elseif ($parameters.Count -ne 0) {
          $errors.Add("Experiment function '$($schemaFunction.name)' must not accept parameters.")
        }
      }
    }
  }
}

$mappedColors = @($manifest.theme_experiment.colors.PSObject.Properties.Name)
$definedColors = @($manifest.theme.colors.PSObject.Properties.Name)
foreach ($color in $mappedColors) {
  if ($color -notin $definedColors) {
    $errors.Add("Theme experiment color '$color' is mapped but not defined in theme.colors.")
  }
}

foreach ($packageRoot in @($sourceDir, $companionDir)) {
  foreach ($item in Get-ChildItem -LiteralPath $packageRoot -Recurse -Force) {
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
      $errors.Add("Package sources must not contain symbolic links or reparse points: $($item.FullName)")
    }
    if (-not $item.PSIsContainer) {
      $relativePath = $item.FullName.Substring($packageRoot.Length).TrimStart([char[]]@('\', '/')).Replace('\', '/')
      if (
        $relativePath -match '(?i)(^|/)(?:\.env(?:\..*)?|\.git(?:/|$)|\.npmrc$|\.pypirc$|id_(?:rsa|dsa|ecdsa|ed25519)(?:\.pub)?$)' -or
        $relativePath -match '(?i)\.(?:key|pem|pfx|p12|jks|keystore|mobileprovision)$'
      ) {
        $errors.Add("Potential secret or private credential must not be packaged: $relativePath")
      }
    }
  }
}

$allowedThemeFiles = [System.Collections.Generic.HashSet[string]]::new(
  [System.StringComparer]::Ordinal
)
$allowedCompanionFiles = [System.Collections.Generic.HashSet[string]]::new(
  [System.StringComparer]::Ordinal
)
foreach ($baseFile in @("LICENSE", "manifest.json")) {
  $allowedThemeFiles.Add($baseFile) | Out-Null
  $allowedCompanionFiles.Add($baseFile) | Out-Null
}
if ($themeExperimentStyle) {
  $allowedThemeFiles.Add($themeExperimentStyle.Replace('\', '/')) | Out-Null
}
foreach ($icon in $manifest.icons.PSObject.Properties.Value) {
  $allowedThemeFiles.Add(([string]$icon).Replace('\', '/')) | Out-Null
}
foreach ($icon in $companionManifest.icons.PSObject.Properties.Value) {
  $allowedCompanionFiles.Add(([string]$icon).Replace('\', '/')) | Out-Null
}
foreach ($backgroundScript in $backgroundScripts) {
  $allowedCompanionFiles.Add(([string]$backgroundScript).Replace('\', '/')) | Out-Null
}
foreach ($experimentProperty in $experimentProperties) {
  $allowedCompanionFiles.Add(([string]$experimentProperty.Value.schema).Replace('\', '/')) | Out-Null
  $allowedCompanionFiles.Add(([string]$experimentProperty.Value.parent.script).Replace('\', '/')) | Out-Null
}

foreach ($packageRecord in @(
  @{ Name = "Theme"; Root = $sourceDir; Allowed = $allowedThemeFiles },
  @{ Name = "Companion"; Root = $companionDir; Allowed = $allowedCompanionFiles }
)) {
  foreach ($file in Get-ChildItem -LiteralPath $packageRecord.Root -Recurse -Force -File) {
    $relativePath = $file.FullName.Substring($packageRecord.Root.Length).TrimStart([char[]]@('\', '/')).Replace('\', '/')
    if (-not $packageRecord.Allowed.Contains($relativePath)) {
      $errors.Add("$($packageRecord.Name) contains an unreviewed package file: $relativePath")
    }
  }
}

$rootLicensePath = Join-Path $projectRoot "LICENSE"
foreach ($packageLicensePath in @(
  (Join-Path $sourceDir "LICENSE"),
  (Join-Path $companionDir "LICENSE")
)) {
  if (-not (Test-Path -LiteralPath $packageLicensePath -PathType Leaf)) {
    $errors.Add("Every package must include LICENSE: $packageLicensePath")
  } elseif (
    (Get-FileHash -Algorithm SHA256 -LiteralPath $packageLicensePath).Hash -ne
    (Get-FileHash -Algorithm SHA256 -LiteralPath $rootLicensePath).Hash
  ) {
    $errors.Add("Packaged LICENSE must exactly match the repository LICENSE: $packageLicensePath")
  }
}

$style = Get-Content -LiteralPath $stylePath -Raw
$unsafeCssPattern = '(?i)javascript\s*:|@import\b|url\s*\(\s*["'']?\s*(?:https?:|//|data\s*:text/html)'
if ($style -match $unsafeCssPattern) {
  $errors.Add("Theme CSS must not contain JavaScript URLs, imports, or remote resources.")
}

$unsafeScriptPattern = '(?i)\beval\s*\(|\bnew\s+Function\s*\(|\.(?:innerHTML|outerHTML)\s*=|\.insertAdjacentHTML\s*\('
$networkScriptPattern = '(?i)\b(?:fetch|XMLHttpRequest|WebSocket|EventSource|importScripts)\b|\.sendBeacon\s*\('
foreach ($scriptFile in Get-ChildItem -LiteralPath $companionDir -Recurse -Force -File -Filter "*.js") {
  $scriptSource = Get-Content -LiteralPath $scriptFile.FullName -Raw
  if ($scriptSource -match $unsafeScriptPattern) {
    $errors.Add("Companion JavaScript must not use dynamic code evaluation or HTML string injection: $($scriptFile.FullName)")
  }
  if ($scriptSource -match $networkScriptPattern) {
    $errors.Add("Companion JavaScript must remain local-only and must not add network clients: $($scriptFile.FullName)")
  }
  $scriptWithoutNamespaceUri = $scriptSource.Replace("http://www.w3.org/1999/xhtml", "")
  if ($scriptWithoutNamespaceUri -match '(?i)https?://') {
    $errors.Add("Companion JavaScript must not embed remote endpoint URLs: $($scriptFile.FullName)")
  }
}

$unsafeSvgPattern = '(?i)<script\b|\bon[a-z]+\s*=|\b(?:href|src)\s*=\s*["'']\s*(?:https?:|//|javascript:|data:)'
foreach ($svgFile in Get-ChildItem -LiteralPath $sourceDir, $companionDir -Recurse -Force -File -Filter "*.svg") {
  if ((Get-Content -LiteralPath $svgFile.FullName -Raw) -match $unsafeSvgPattern) {
    $errors.Add("SVG assets must not contain scripts, event handlers, or external data sources: $($svgFile.FullName)")
  }
}

if ($errors.Count -gt 0) {
  $errors | ForEach-Object { Write-Error $_ }
  throw "Validation failed with $($errors.Count) error(s)."
}

Write-Output "Validated metadata, compatibility, package paths, least privilege, licenses, and local-only assets."
