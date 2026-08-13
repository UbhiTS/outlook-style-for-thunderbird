[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $projectRoot "src"
$companionDir = Join-Path $projectRoot "companion"
$manifestPath = Join-Path $sourceDir "manifest.json"
$stylePath = Join-Path $sourceDir "styles\outlook-fluent.css"
$companionManifestPath = Join-Path $companionDir "manifest.json"

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$companionManifest = Get-Content -LiteralPath $companionManifestPath -Raw | ConvertFrom-Json

$errors = [System.Collections.Generic.List[string]]::new()

if ($manifest.manifest_version -ne 3) {
  $errors.Add("manifest_version must be 3.")
}
if (-not $manifest.browser_specific_settings.gecko.id) {
  $errors.Add("browser_specific_settings.gecko.id is required for permanent Thunderbird installation.")
}
if ($manifest.browser_specific_settings.gecko.strict_min_version -ne "153.0.3") {
  $errors.Add("strict_min_version must be 153.0.3.")
}
if (-not $manifest.theme.colors.frame -or -not $manifest.theme.colors.tab_background_text) {
  $errors.Add("The static theme must define frame and tab_background_text colors.")
}
if ($manifest.background -or $manifest.permissions -or $manifest.host_permissions) {
  $errors.Add("This static theme must not request scripts or permissions.")
}
if ($companionManifest.version -ne $manifest.version) {
  $errors.Add("Theme and companion versions must match.")
}
if (
  @($companionManifest.permissions) -notcontains "messagesRead" -or
  @($companionManifest.permissions) -notcontains "scripting"
) {
  $errors.Add("The companion requires messagesRead and scripting for thread selection and the message-display API.")
}
if (-not $companionManifest.background.scripts) {
  $errors.Add("The companion background registration script is required.")
}
foreach ($messageDisplayScript in @($companionManifest.message_display_scripts)) {
  foreach ($cssPath in @($messageDisplayScript.css)) {
    if (-not (Test-Path -LiteralPath (Join-Path $companionDir ([string]$cssPath)))) {
      $errors.Add("Referenced companion stylesheet does not exist: $cssPath")
    }
  }
  foreach ($jsPath in @($messageDisplayScript.js)) {
    if (-not (Test-Path -LiteralPath (Join-Path $companionDir ([string]$jsPath)))) {
      $errors.Add("Referenced companion script does not exist: $jsPath")
    }
  }
}
foreach ($backgroundScript in @($companionManifest.background.scripts)) {
  if (-not (Test-Path -LiteralPath (Join-Path $companionDir ([string]$backgroundScript)))) {
    $errors.Add("Referenced companion background script does not exist: $backgroundScript")
  }
}
foreach ($experiment in $companionManifest.experiment_apis.PSObject.Properties.Value) {
  foreach ($experimentPath in @($experiment.schema, $experiment.parent.script)) {
    if (-not (Test-Path -LiteralPath (Join-Path $companionDir ([string]$experimentPath)))) {
      $errors.Add("Referenced companion Experiment file does not exist: $experimentPath")
    }
  }
}

$themeExperimentStyle = [string]$manifest.theme_experiment.stylesheet
if (-not $themeExperimentStyle) {
  $errors.Add("theme_experiment.stylesheet is required for the detailed Outlook styling.")
} elseif (-not (Test-Path -LiteralPath (Join-Path $sourceDir $themeExperimentStyle))) {
  $errors.Add("Referenced theme stylesheet does not exist: $themeExperimentStyle")
}

foreach ($icon in $manifest.icons.PSObject.Properties.Value) {
  if (-not (Test-Path -LiteralPath (Join-Path $sourceDir ([string]$icon)))) {
    $errors.Add("Referenced icon does not exist: $icon")
  }
}

$mappedColors = @($manifest.theme_experiment.colors.PSObject.Properties.Name)
$definedColors = @($manifest.theme.colors.PSObject.Properties.Name)
foreach ($color in $mappedColors) {
  if ($color -notin $definedColors) {
    $errors.Add("Theme experiment color '$color' is mapped but not defined in theme.colors.")
  }
}

$style = Get-Content -LiteralPath $stylePath -Raw
$remotePattern = '(?i)javascript\s*:|@import\s+url\s*\(\s*["'']?https?://'
if ($style -match $remotePattern) {
  $errors.Add("Theme CSS must not contain JavaScript URLs or remote imports.")
}

if ($errors.Count -gt 0) {
  $errors | ForEach-Object { Write-Error $_ }
  throw "Validation failed with $($errors.Count) error(s)."
}

Write-Output "Manifest, references, permissions, and theme color mappings are valid."
