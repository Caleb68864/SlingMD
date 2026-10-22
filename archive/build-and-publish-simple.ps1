param(
    [switch]$SkipPublish
)

Write-Host "===== SlingMD Build & Publish Script =====" -ForegroundColor Cyan

# First build the solution
Write-Host "Building solution..." -ForegroundColor Cyan
dotnet build SlingMD.sln --configuration Release
$buildSucceeded = $?

if (-not $buildSucceeded) {
    Write-Host "BUILD FAILED" -ForegroundColor Red
    exit 1
}

Write-Host "Build completed successfully" -ForegroundColor Green

# If requested to skip publish, exit here
if ($SkipPublish) {
    Write-Host "Skipping publish as requested" -ForegroundColor Yellow
    exit 0
}

# Publish the project
Write-Host "Publishing project..." -ForegroundColor Cyan
dotnet publish SlingMD.Outlook\SlingMD.Outlook.csproj --configuration Release
$publishSucceeded = $?

if (-not $publishSucceeded) {
    Write-Host "PUBLISH FAILED" -ForegroundColor Red
    exit 1
}

Write-Host "Publish completed successfully" -ForegroundColor Green

# Run packaging script if it exists
if (Test-Path ".\package-release.ps1") {
    Write-Host "Running package script..." -ForegroundColor Cyan
    .\package-release.ps1
    $packageSucceeded = $?
    
    if (-not $packageSucceeded) {
        Write-Host "PACKAGING FAILED" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Packaging completed successfully" -ForegroundColor Green
}

Write-Host "===== All operations completed successfully =====" -ForegroundColor Green
exit 0 