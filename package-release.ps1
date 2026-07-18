# Packages the latest ClickOnce publish output into a versioned ZIP under Releases\.
#
# IMPORTANT (issue #13 postmortem): a previous release shipped a STALE compiled DLL because
# this script only zips whatever already sits in publish\Application Files\ -- it does not
# build or publish. If the ClickOnce output was not re-published after the last source change,
# the ZIP silently ships pre-fix binaries. To make that impossible, this script now:
#   1. Builds the add-in in Release (so bin\Release is the current source of truth).
#   2. Selects the newest publish folder.
#   3. Refuses to package unless the published DLL byte-for-byte matches the freshly built DLL.
# If they differ, you must (re)Publish from Visual Studio (Build > Publish) before packaging.

$ErrorActionPreference = "Stop"

# --- Resolve tooling ---------------------------------------------------------
$msbuild = $null
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
}
if (-not $msbuild -or -not (Test-Path $msbuild)) {
    Write-Error "Could not locate MSBuild via vswhere. Build+Publish from Visual Studio, then rerun."
    exit 1
}

# --- 1. Build Release (current source of truth) ------------------------------
# IMPORTANT: this script does NOT publish. Regenerating the ClickOnce payload
# (publish\Application Files\<version>\*.deploy) requires Visual Studio's VSTO publish
# pipeline -- "Build > Publish SlingMD.Outlook". The CLI MSBuild Publish target does NOT run
# that pipeline for Office/VSTO projects (it only builds bin\Release), so it cannot refresh the
# payload. RUN VS PUBLISH FIRST, then run this script. The SHA guard in step 3 enforces that:
# it refuses to package a payload that doesn't match the freshly built DLL below.
Write-Host "Building SlingMD.Outlook (Release)..."
& $msbuild ".\SlingMD.Outlook\SlingMD.Outlook.csproj" -t:Build -p:Configuration=Release -v:m -nologo
if ($LASTEXITCODE -ne 0) {
    Write-Error "Release build failed. Fix build errors before packaging."
    exit 1
}
$builtDll = ".\SlingMD.Outlook\bin\Release\SlingMD.Outlook.dll"
if (-not (Test-Path $builtDll)) {
    Write-Error "Expected built DLL not found: $builtDll"
    exit 1
}
$builtHash = (Get-FileHash $builtDll -Algorithm SHA256).Hash

# --- 2. Select newest publish folder -----------------------------------------
$appFilesPath = ".\SlingMD.Outlook\publish\Application Files"
$latestVersion = Get-ChildItem $appFilesPath |
    Where-Object { $_.Name -like "SlingMD.Outlook_*" } |
    Sort-Object -Property {
        $version = $_.Name -replace "SlingMD\.Outlook_", "" -replace "_", "."
        [version]$version
    } -Descending |
    Select-Object -First 1

if (-not $latestVersion) {
    Write-Error "No version found in Application Files directory"
    exit 1
}

# --- 3. Staleness guard: published DLL must match freshly built DLL -----------
$publishedDll = Join-Path $latestVersion.FullName "SlingMD.Outlook.dll.deploy"
if (-not (Test-Path $publishedDll)) {
    $publishedDll = Join-Path $latestVersion.FullName "SlingMD.Outlook.dll"
}
if (-not (Test-Path $publishedDll)) {
    Write-Error "No published SlingMD.Outlook.dll(.deploy) found in $($latestVersion.Name)."
    exit 1
}
$publishedHash = (Get-FileHash $publishedDll -Algorithm SHA256).Hash

if ($publishedHash -ne $builtHash) {
    Write-Error @"
STALE PUBLISH DETECTED -- refusing to package (this is the issue #13 failure mode).

  Newest publish folder : $($latestVersion.Name)
  Published DLL SHA256   : $publishedHash
  Freshly built SHA256   : $builtHash

The ClickOnce publish output does NOT match current source, so packaging it would ship
pre-fix binaries. Re-publish first:
  Visual Studio > right-click SlingMD.Outlook > Publish (or Build > Publish SlingMD.Outlook),
then rerun .\package-release.ps1.
"@
    exit 1
}

# Extract version number
$versionNumber = $latestVersion.Name -replace "SlingMD\.Outlook_", "" -replace "_", "."

# Create Releases directory if it doesn't exist
$releasesDir = ".\Releases"
if (-not (Test-Path $releasesDir)) {
    New-Item -ItemType Directory -Path $releasesDir | Out-Null
}

# Create zip file name
$zipFileName = "SlingMD.Outlook_$($versionNumber -replace "\.", "_").zip"
$zipFilePath = Join-Path $releasesDir $zipFileName

# Create a temporary directory for the files we want to include
$tempDir = ".\temp_package"
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempDir | Out-Null

# Copy only the necessary files
Copy-Item ".\SlingMD.Outlook\publish\SlingMD.Outlook.vsto" $tempDir
Copy-Item ".\SlingMD.Outlook\publish\setup.exe" $tempDir
Copy-Item $latestVersion.FullName "$tempDir\Application Files\$($latestVersion.Name)" -Recurse

# Create the zip file
Compress-Archive -Path "$tempDir\*" -DestinationPath $zipFilePath -Force

# Clean up temp directory
Remove-Item $tempDir -Recurse -Force

Write-Host "Package created successfully: $zipFileName (verified matches Release build $builtHash)"
