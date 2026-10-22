param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipPublish = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipClean = $false
)

# Set working directory to script location
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# Function to write colored output with timestamp
function Write-Log {
    param(
        [string]$Message,
        [System.ConsoleColor]$Color = [System.ConsoleColor]::White
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$timestamp] " -NoNewline -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor $Color
}

# Function to detect if project is VSTO
function Test-IsVSTOProject {
    param(
        [string]$ProjectPath
    )

    if (Test-Path $ProjectPath) {
        $content = Get-Content $ProjectPath -Raw
        return $content -match "OfficeTools\\Microsoft\.VisualStudio\.Tools\.Office\.targets" -or
               $content -match "ProjectTypeGuids.*BAA0C2D2-18E2-41B9-852F-F413020CAA33"
    }
    return $false
}

# Function to validate Office Tools installation
function Test-OfficeTools {
    param(
        [string]$VSPath
    )

    if (-not $VSPath) {
        return $false
    }

    # Check for Office Tools targets file
    $officeToolsPath = Join-Path $VSPath "MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"

    if (Test-Path $officeToolsPath) {
        Write-Log "Found Office/SharePoint development tools" -Color Green
        return $true
    }
    else {
        Write-Log "Office/SharePoint development tools NOT found" -Color Yellow
        return $false
    }
}

# Function to find build tool
function Find-BuildTool {
    Write-Log "Searching for build tools..." -Color Gray

    # Check if this is a VSTO project
    $isVSTO = Test-IsVSTOProject "SlingMD.Outlook\SlingMD.Outlook.csproj"

    if ($isVSTO) {
        Write-Log "VSTO project detected - MSBuild with Office Tools required" -Color Cyan

        # For VSTO projects, we MUST use MSBuild with Office Tools
        try {
            $vsPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
            if ($vsPath) {
                $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
                if (Test-Path $msbuildPath) {
                    Write-Log "Found MSBuild at: $msbuildPath" -Color Green

                    # Validate Office Tools are installed
                    $hasOfficeTools = Test-OfficeTools -VSPath $vsPath

                    if (-not $hasOfficeTools) {
                        Write-Log "" -Color Red
                        Write-Log "=====================================================================" -Color Red
                        Write-Log "ERROR: Office/SharePoint Development Tools NOT INSTALLED" -Color Red
                        Write-Log "=====================================================================" -Color Red
                        Write-Log "This is a VSTO project that requires Office development tools." -Color Yellow
                        Write-Log "" -Color Yellow
                        Write-Log "To fix this:" -Color Yellow
                        Write-Log "1. Run Visual Studio Installer" -Color White
                        Write-Log "2. Click 'Modify' on your Visual Studio installation" -Color White
                        Write-Log "3. Select 'Office/SharePoint development' workload" -Color White
                        Write-Log "4. Click 'Modify' to install" -Color White
                        Write-Log "" -Color Yellow
                        Write-Log "Current Visual Studio: $vsPath" -Color Gray
                        Write-Log "=====================================================================" -Color Red
                        return $null
                    }

                    return $msbuildPath
                }
            }
        }
        catch {
            Write-Log "Error searching for Visual Studio: $($_.Exception.Message)" -Color Red
        }

        Write-Log "" -Color Red
        Write-Log "=====================================================================" -Color Red
        Write-Log "ERROR: Visual Studio with Office Tools NOT FOUND" -Color Red
        Write-Log "=====================================================================" -Color Red
        Write-Log "VSTO projects require Visual Studio with Office/SharePoint tools." -Color Yellow
        Write-Log "" -Color Yellow
        Write-Log "Please install one of:" -Color Yellow
        Write-Log "- Visual Studio 2022 Community (FREE)" -Color White
        Write-Log "- Visual Studio 2022 Professional" -Color White
        Write-Log "- Visual Studio 2022 Enterprise" -Color White
        Write-Log "" -Color Yellow
        Write-Log "With the 'Office/SharePoint development' workload selected." -Color Yellow
        Write-Log "=====================================================================" -Color Red
        return $null
    }

    # For non-VSTO projects, try dotnet CLI first
    if (Get-Command "dotnet" -ErrorAction SilentlyContinue) {
        Write-Log "Found dotnet CLI" -Color Green
        return "dotnet"
    }

    # Try MSBuild through Visual Studio
    try {
        $vsPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsPath) {
            $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msbuildPath) {
                Write-Log "Found MSBuild at: $msbuildPath" -Color Green
                return $msbuildPath
            }
        }
    }
    catch {
        # Continue to next option
    }

    Write-Log "No build tools found. Please install .NET SDK or Visual Studio." -Color Red
    return $null
}

# Function to clean the solution
function Start-Clean {
    param(
        [string]$BuildTool
    )

    if ($SkipClean) {
        Write-Log "Skipping solution cleaning as requested" -Color Yellow
        return $true
    }

    Write-Log "Cleaning solution..." -Color Gray

    if ($BuildTool -eq "dotnet") {
        dotnet clean SlingMD.sln --configuration $Configuration --verbosity minimal 2>&1 | Out-Host
    }
    else {
        & $BuildTool SlingMD.sln /t:Clean /p:Configuration=$Configuration /v:minimal /nologo 2>&1 | Out-Host
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Log "Clean operation failed with exit code $LASTEXITCODE" -Color Red
        Write-Log "This may be OK - continuing with build..." -Color Yellow
        return $true  # Don't fail on clean errors
    }

    Write-Log "Clean operation completed successfully" -Color Green
    return $true
}

# Function to build the solution
function Start-Build {
    param(
        [string]$BuildTool
    )

    Write-Log "Building solution..." -Color Gray

    if ($BuildTool -eq "dotnet") {
        dotnet build SlingMD.sln --configuration $Configuration --verbosity minimal 2>&1 | Out-Host
    }
    else {
        & $BuildTool SlingMD.sln /t:Build /p:Configuration=$Configuration /v:minimal /nologo 2>&1 | Out-Host
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Log "" -Color Red
        Write-Log "Build failed with exit code $LASTEXITCODE" -Color Red
        Write-Log "Check the error messages above for details." -Color Yellow
        return $false
    }

    Write-Log "Build completed successfully" -Color Green
    return $true
}

# Function to publish the project
function Start-Publish {
    param(
        [string]$BuildTool
    )

    Write-Log "Publishing SlingMD..." -Color Gray

    if ($BuildTool -eq "dotnet") {
        dotnet publish SlingMD.Outlook\SlingMD.Outlook.csproj --configuration $Configuration --verbosity minimal 2>&1 | Out-Host
    }
    else {
        & $BuildTool SlingMD.Outlook\SlingMD.Outlook.csproj /t:Publish /p:Configuration=$Configuration /v:minimal /nologo 2>&1 | Out-Host
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Log "" -Color Red
        Write-Log "Publish operation failed with exit code $LASTEXITCODE" -Color Red
        Write-Log "Check the error messages above for details." -Color Yellow
        return $false
    }

    # Run package script if it exists
    if (Test-Path ".\package-release.ps1") {
        Write-Log "Running package script..." -Color Gray
        & ".\package-release.ps1" 2>&1 | Out-Host

        if ($LASTEXITCODE -ne 0) {
            Write-Log "Package script failed with exit code $LASTEXITCODE" -Color Yellow
            Write-Log "Continuing anyway..." -Color Yellow
        }
    }

    Write-Log "Publish completed successfully" -Color Green
    return $true
}

# Main execution
Write-Log "SlingMD Build and Publish Tool" -Color Magenta
Write-Log "----------------------------" -Color Magenta
Write-Log "Configuration: $Configuration" -Color Gray
Write-Log "Skip Clean: $SkipClean" -Color Gray
Write-Log "Skip Publish: $SkipPublish" -Color Gray
Write-Log "----------------------------" -Color Magenta

# Find build tool
$buildTool = Find-BuildTool
if (-not $buildTool) {
    exit 1
}

# Clean solution (don't fail if clean has issues)
$cleanSuccess = Start-Clean -BuildTool $buildTool

# Build solution
$buildSuccess = Start-Build -BuildTool $buildTool
if (-not $buildSuccess) {
    Write-Log "Build failed. Publish operation skipped." -Color Red
    exit 1
}

# Publish if needed
if (-not $SkipPublish) {
    $publishSuccess = Start-Publish -BuildTool $buildTool
    if (-not $publishSuccess) {
        Write-Log "Publish operation failed." -Color Red
        exit 1
    }
    
    Write-Log "SlingMD was successfully built and published!" -Color Green
} 
else {
    Write-Log "SlingMD was successfully built. Publish was skipped." -Color Green
}

exit 0 