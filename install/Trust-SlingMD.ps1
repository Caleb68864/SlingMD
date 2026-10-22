# Prepares this computer to install SlingMD, then offers to run setup.exe.
#
# SlingMD is signed with its own self-signed certificate ("SlingMD Development"), not one from a
# commercial certificate authority. Windows and Office don't know that certificate, so an
# installer extracted from a downloaded ZIP can fail with "publisher cannot be verified" or
# "not signed". This script, run from the extracted release folder:
#   1. Removes the "downloaded from the internet" mark from every file in this folder.
#   2. Adds SlingMD.cer (the public half of the signing certificate) to this computer's
#      Trusted Root and Trusted Publishers stores, so SlingMD and its updates install without
#      the trust prompt.
#   3. Offers to run setup.exe.
#
# Needs administrator rights for step 2; it asks for them when started without.
# Easiest way to run it: double-click Trust-SlingMD.cmd in the same folder.

param(
    [switch]$NoInstall
)

$ErrorActionPreference = "Stop"
$folder = Split-Path -Parent $MyInvocation.MyCommand.Path

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Administrator rights are needed to trust the certificate. Asking for them..."
    $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$($MyInvocation.MyCommand.Path)`"")
    if ($NoInstall) { $arguments += "-NoInstall" }
    Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
    exit
}

try {
    # 1. Unblock
    Get-ChildItem -LiteralPath $folder -Recurse -File | Unblock-File
    Write-Host "Unblocked the files in $folder"

    # 2. Trust the certificate
    $certPath = Join-Path $folder "SlingMD.cer"
    if (-not (Test-Path -LiteralPath $certPath)) {
        throw "SlingMD.cer was not found next to this script. Extract the whole ZIP and run it again."
    }

    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
    Write-Host ""
    Write-Host "Certificate : $($cert.Subject)"
    Write-Host "Thumbprint  : $($cert.Thumbprint)"
    Write-Host "Valid until : $($cert.NotAfter.ToString('yyyy-MM-dd'))"

    foreach ($store in @("Root", "TrustedPublisher")) {
        Import-Certificate -FilePath $certPath -CertStoreLocation "Cert:\LocalMachine\$store" | Out-Null
        Write-Host "Added to LocalMachine\$store"
    }

    # 3. Install
    $setup = Join-Path $folder "setup.exe"
    Write-Host ""
    if (-not $NoInstall -and (Test-Path -LiteralPath $setup)) {
        $answer = Read-Host "Run setup.exe now? [Y/n]"
        if ($answer -eq "" -or $answer -match "^[Yy]") {
            Start-Process -FilePath $setup -WorkingDirectory $folder
        }
    }
    else {
        Write-Host "Done. Run setup.exe to install SlingMD."
    }
}
catch {
    Write-Host ""
    Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "If your organisation blocks this, ask IT to trust SlingMD.cer for you."
}

Write-Host ""
Read-Host "Press Enter to close"
