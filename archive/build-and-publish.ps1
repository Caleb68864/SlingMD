param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipPublish = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipClean = $false
)

# Set error action preference to stop on any error
$ErrorActionPreference = "Stop"

# Set working directory to script location
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# Define paths
$slnPath = ".\SlingMD.sln"
$outlookProjPath = ".\SlingMD.Outlook\SlingMD.Outlook.csproj"
$publishDir = ".\SlingMD.Outlook\publish"

# Function to write colored output with timestamp
function Write-ColorOutput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [System.ConsoleColor]$ForegroundColor = [System.ConsoleColor]::White
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$timestamp] " -NoNewline -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Ensure MSBuild is in path
function Ensure-MSBuild {
    try {
        # First try direct dotnet command
        if (Get-Command "dotnet" -ErrorAction SilentlyContinue) {
            Write-ColorOutput "Using dotnet CLI for build operations" -ForegroundColor Green
            return "dotnet"
        }
        
        # Try to find MSBuild in usual locations
        $vsPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsPath) {
            $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msbuildPath) {
                return $msbuildPath
            }
        }
        
        # If we can't find it in the default location, try a wildcard search
        $vsPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsPath) {
            $candidates = Get-ChildItem "$vsPath\MSBuild" -Recurse -Filter "MSBuild.exe" | Select-Object -First 1
            if ($candidates) {
                return $candidates.FullName
            }
        }
        
        # If still not found, throw an error
        throw "Neither dotnet CLI nor MSBuild.exe was found. Please ensure .NET SDK or Visual Studio is installed."
    }
    catch {
        Write-ColorOutput "Error finding build tools: $_" -ForegroundColor Red
        exit 1
    }
}

# Main build function
function Build-Project {
    Write-ColorOutput "Building SlingMD in $Configuration configuration..." -ForegroundColor Cyan
    
    try {
        $buildTool = Ensure-MSBuild
        Write-ColorOutput "Using build tool: $buildTool" -ForegroundColor Gray
        
        # Clean the solution if not skipping
        if (-not $SkipClean) {
            Write-ColorOutput "Cleaning solution..." -ForegroundColor Gray
            
            if ($buildTool -eq "dotnet") {
                & dotnet clean $slnPath --configuration $Configuration --verbosity minimal
            }
            else {
                & $buildTool $slnPath /t:Clean /p:Configuration=$Configuration /v:minimal
            }
            
            if ($LASTEXITCODE -ne 0) {
                Write-ColorOutput "Clean failed with exit code $LASTEXITCODE" -ForegroundColor Red
                return $false
            }
            
            Write-ColorOutput "Clean completed successfully." -ForegroundColor Green
        }
        else {
            Write-ColorOutput "Clean operation skipped." -ForegroundColor Yellow
        }
        
        # Build the solution
        Write-ColorOutput "Building solution..." -ForegroundColor Gray
        
        if ($buildTool -eq "dotnet") {
            & dotnet build $slnPath --configuration $Configuration --verbosity minimal
        }
        else {
            & $buildTool $slnPath /t:Build /p:Configuration=$Configuration /v:minimal
        }
        
        if ($LASTEXITCODE -ne 0) {
            Write-ColorOutput "Build failed with exit code $LASTEXITCODE" -ForegroundColor Red
            return $false
        }
        
        Write-ColorOutput "Build completed successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-ColorOutput "Build failed: $_" -ForegroundColor Red
        return $false
    }
}

# Publish function
function Publish-Project {
    Write-ColorOutput "Publishing SlingMD..." -ForegroundColor Cyan
    
    try {
        $buildTool = Ensure-MSBuild
        
        # Publish the project
        Write-ColorOutput "Publishing project..." -ForegroundColor Gray
        
        if ($buildTool -eq "dotnet") {
            & dotnet publish $outlookProjPath --configuration $Configuration --verbosity minimal
        }
        else {
            & $buildTool $outlookProjPath /t:Publish /p:Configuration=$Configuration /v:minimal
        }
        
        if ($LASTEXITCODE -ne 0) {
            Write-ColorOutput "Publish failed with exit code $LASTEXITCODE" -ForegroundColor Red
            return $false
        }
        
        # Run the existing package script if it exists
        if (Test-Path ".\package-release.ps1") {
            Write-ColorOutput "Running package-release.ps1 script..." -ForegroundColor Gray
            & ".\package-release.ps1"
            
            if ($LASTEXITCODE -ne 0) {
                Write-ColorOutput "Package script failed with exit code $LASTEXITCODE" -ForegroundColor Red
                return $false
            }
            
            Write-ColorOutput "Package script completed successfully." -ForegroundColor Green
        }
        else {
            Write-ColorOutput "No package-release.ps1 script found, skipping packaging step." -ForegroundColor Yellow
        }
        
        Write-ColorOutput "Publish completed successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-ColorOutput "Publish failed: $_" -ForegroundColor Red
        return $false
    }
}

# Main script execution
try {
    # Display script info
    Write-ColorOutput "SlingMD Build and Publish Script" -ForegroundColor Magenta
    Write-ColorOutput "--------------------------------" -ForegroundColor Magenta
    Write-ColorOutput "Configuration: $Configuration" -ForegroundColor Gray
    Write-ColorOutput "Skip Clean: $SkipClean" -ForegroundColor Gray
    Write-ColorOutput "Skip Publish: $SkipPublish" -ForegroundColor Gray
    Write-ColorOutput "--------------------------------" -ForegroundColor Magenta
    
    # Build the project
    $buildSuccess = Build-Project
    
    # Debug the build result
    Write-ColorOutput "Build result: $buildSuccess" -ForegroundColor Yellow
    
    # If build succeeded and not skipping publish, publish the project
    if ($buildSuccess) {
        if (-not $SkipPublish) {
            $publishSuccess = Publish-Project
            
            if ($publishSuccess) {
                Write-ColorOutput "SlingMD was successfully built and published!" -ForegroundColor Green
                exit 0
            }
            else {
                Write-ColorOutput "SlingMD was built but publishing failed." -ForegroundColor Yellow
                exit 1
            }
        }
        else {
            Write-ColorOutput "SlingMD was successfully built. Publish was skipped." -ForegroundColor Green
            exit 0
        }
    }
    else {
        Write-ColorOutput "SlingMD build failed. Publish was skipped." -ForegroundColor Red
        exit 1
    }
}
catch {
    Write-ColorOutput "Error: $_" -ForegroundColor Red
    exit 1
} 