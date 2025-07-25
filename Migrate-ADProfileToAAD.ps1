# Requires -RunAsAdministrator
# Modules: need MSOnline or AzureAD for extended functionality
function Test-WindowsEdition {
    $edition = (Get-WmiObject -Class Win32_OperatingSystem).OperatingSystemSKU
    if ($edition -ne 48 -and $edition -ne 4) {
        Write-Error "This script requires Windows 10/11 Pro or Enterprise. Current SKU: $edition"
        exit 1
    }
}


function Save-WindowsCredentials {
    param ($SourceProfile)
    $credPath = "$env:ProgramData\AADMigration\$SourceProfile-WinCreds.xml"
    cmdkey /list | Export-Clixml -Path $credPath
    Write-Output "Credentials saved to $credPath"
    return $credPath
}

function Import-WindowsCredentials {
    param (
        [Parameter(Mandatory = $true)]
        [System.Security.SecureString]$CredentialFilePath
    )
    if (Test-Path $CredentialFilePath) {
        $credentialFilePathString = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CredentialFilePath))
        $creds = Import-Clixml -Path $credentialFilePathString
        foreach ($cred in $creds) {
            cmdkey /add:$cred.Target /user:$cred.UserName /pass:$cred.Password
        }
        Write-Output "Credentials imported."
    } else {
        Write-Warning "Credential file not found."
    }
}

function Repair-OneDrivePath {
    param (
        [string]$TargetProfile
    )

    $newProfilePath = "C:\Users\$TargetProfile"
    $oldOneDrivePath = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\OneDrive" -ErrorAction SilentlyContinue).UserFolder

    if ($oldOneDrivePath -and ($oldOneDrivePath -notlike "$newProfilePath*")) {
        Write-Warning "OneDrive is still pointing to the old profile path: $oldOneDrivePath"
        Write-Output "Resetting OneDrive configuration..."

        # Stop OneDrive
        Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue

        # Reset OneDrive
        Start-Process -FilePath "$env:SystemRoot\System32\OneDriveSetup.exe" -ArgumentList "/reset" -Wait

        # Start OneDrive again
        Start-Process -FilePath "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe"

        Write-Output "OneDrive has been reset. Please sign in again to configure the correct sync path."
    } else {
        Write-Output "OneDrive path is already correct or not configured yet."
    }
}

function Start-MDMSync {
    Write-Output "Triggering MDM sync..."
    Start-Process -FilePath "dsregcmd.exe" -ArgumentList "/status"
    Start-Process -FilePath "dsregcmd.exe" -ArgumentList "/sync"
}

function Start-ProfileMigration {
    Test-WindowsEdition

    $sourceProfile = Read-Host "Enter the source (local AD) profile name"
    $targetProfile = Read-Host "Enter the target (AAD) profile name"

    $savedCredFile = Save-WindowsCredentials -SourceProfile $sourceProfile
    Write-Output "Credentials saved to: $savedCredFile"
    Write-Output "Please reboot and sign in with the AAD account ($targetProfile). Then re-run this script with the -PostReboot flag."
}

function Complete-PostRebootTasks {
    $targetProfile = Read-Host "Enter the target (AAD) profile name"
    $credentialFile = "$env:ProgramData\AADMigration\$targetProfile-WinCreds.xml"

    Import-WindowsCredentials -CredentialFilePath $credentialFile
    Repair-OneDrivePath -TargetProfile $targetProfile
    Start-MDMSync

    Write-Output "Please sign into Office apps manually to complete the setup."
}

param (
    [switch]$PostReboot
)

if ($PostReboot) {
    Complete-PostRebootTasks
} else {
    Start-ProfileMigration
}
