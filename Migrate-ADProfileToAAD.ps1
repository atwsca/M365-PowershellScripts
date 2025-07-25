<#
.SYNOPSIS
    Migrates user profiles from on-premises Active Directory to Azure Active Directory (AAD).

.DESCRIPTION
    This script facilitates the migration of user profiles from local Active Directory to Azure Active Directory.
    It handles credential preservation, OneDrive path correction, and MDM synchronization during the migration process.
    The script operates in two phases: pre-reboot preparation and post-reboot completion.

.PARAMETER PostReboot
    Switch parameter that indicates the script should run post-reboot tasks instead of initial migration preparation.

.EXAMPLE
    .\Migrate-ADProfileToAAD.ps1
    Runs the initial migration preparation phase, saving credentials and preparing for reboot.

.EXAMPLE
    .\Migrate-ADProfileToAAD.ps1 -PostReboot
    Runs the post-reboot completion phase, restoring credentials and configuring services.

.NOTES
    File Name      : Migrate-ADProfileToAAD.ps1
    Author         : ATWS
    Prerequisite   : PowerShell 5.1 or later, Windows 10/11 Pro or Enterprise
    Copyright      : (c) ATWS. All rights reserved.

.LINK
    https://github.com/atwsca/M365-PoweshellScripts

#>

# Requires -RunAsAdministrator
# Modules: need MSOnline or AzureAD for extended functionality

<#
.SYNOPSIS
    Tests if the current Windows edition supports the migration process.

.DESCRIPTION
    Verifies that the current Windows operating system is either Windows 10/11 Pro or Enterprise,
    as these editions are required for proper AAD integration and domain join capabilities.

.EXAMPLE
    Test-WindowsEdition
    Validates the Windows edition and exits with error code 1 if unsupported.

.NOTES
    Supported SKUs: 48 (Professional), 4 (Enterprise)
#>
function Test-WindowsEdition {
    $edition = (Get-WmiObject -Class Win32_OperatingSystem).OperatingSystemSKU
    if ($edition -ne 48 -and $edition -ne 4) {
        Write-Error "This script requires Windows 10/11 Pro or Enterprise. Current SKU: $edition"
        exit 1
    }
}


<#
.SYNOPSIS
    Saves Windows credentials for the specified source profile.

.DESCRIPTION
    Exports the current user's stored Windows credentials to an XML file for later restoration
    after the profile migration is complete. This ensures that saved passwords and certificates
    are preserved during the migration process.

.PARAMETER SourceProfile
    The name of the source (local AD) user profile whose credentials should be saved.

.EXAMPLE
    Save-WindowsCredentials -SourceProfile "john.doe"
    Saves credentials for the john.doe profile to the migration directory.

.OUTPUTS
    System.String
    Returns the path to the saved credentials file.

.NOTES
    Credentials are saved to $env:ProgramData\AADMigration\{SourceProfile}-WinCreds.xml
#>
function Save-WindowsCredentials {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceProfile
    )
    $credPath = "$env:ProgramData\AADMigration\$SourceProfile-WinCreds.xml"
    cmdkey /list | Export-Clixml -Path $credPath
    Write-Output "Credentials saved to $credPath"
    return $credPath
}

<#
.SYNOPSIS
    Imports previously saved Windows credentials.

.DESCRIPTION
    Restores Windows credentials that were saved during the pre-reboot phase of the migration.
    This function reads the credential XML file and re-adds the stored credentials to the
    Windows Credential Manager.

.PARAMETER CredentialFilePath
    Secure string containing the path to the credential file to import.

.EXAMPLE
    $secureFilePath = ConvertTo-SecureString "C:\ProgramData\AADMigration\user-WinCreds.xml" -AsPlainText -Force
    Import-WindowsCredentials -CredentialFilePath $secureFilePath
    Imports credentials from the specified secure file path.

.NOTES
    The credential file must exist and be in the correct XML format created by Save-WindowsCredentials.
#>
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

<#
.SYNOPSIS
    Repairs OneDrive synchronization path after profile migration.

.DESCRIPTION
    Checks if OneDrive is still pointing to the old profile path and resets the OneDrive
    configuration if necessary. This ensures that OneDrive syncs to the correct location
    for the new AAD profile.

.PARAMETER TargetProfile
    The name of the target (AAD) user profile that OneDrive should sync to.

.EXAMPLE
    Repair-OneDrivePath -TargetProfile "john.doe@company.com"
    Repairs OneDrive path for the specified AAD profile.

.NOTES
    This function will stop OneDrive, reset its configuration, and restart it if the path is incorrect.
    The user will need to sign in to OneDrive again after the reset.
#>
function Repair-OneDrivePath {
    param (
        [Parameter(Mandatory = $true)]
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

<#
.SYNOPSIS
    Initiates Mobile Device Management (MDM) synchronization.

.DESCRIPTION
    Triggers MDM sync to ensure that the device is properly registered and configured
    with Azure AD and any associated MDM policies are applied.

.EXAMPLE
    Start-MDMSync
    Triggers MDM synchronization and displays device registration status.

.NOTES
    Uses dsregcmd.exe to check status and force synchronization with Azure AD.
#>
function Start-MDMSync {
    Write-Output "Triggering MDM sync..."
    Start-Process -FilePath "dsregcmd.exe" -ArgumentList "/status"
    Start-Process -FilePath "dsregcmd.exe" -ArgumentList "/sync"
}

<#
.SYNOPSIS
    Initiates the profile migration process (pre-reboot phase).

.DESCRIPTION
    Starts the first phase of profile migration by validating the Windows edition,
    collecting source and target profile information, saving credentials, and
    preparing the system for reboot with the new AAD account.

.EXAMPLE
    Start-ProfileMigration
    Begins the profile migration process with interactive prompts for profile names.

.NOTES
    This function requires administrator privileges and will prompt for source and target profile names.
    After completion, the system should be rebooted and signed in with the AAD account.
#>
function Start-ProfileMigration {
    Test-WindowsEdition

    $sourceProfile = Read-Host "Enter the source (local AD) profile name"
    $targetProfile = Read-Host "Enter the target (AAD) profile name"

    $savedCredFile = Save-WindowsCredentials -SourceProfile $sourceProfile
    Write-Output "Credentials saved to: $savedCredFile"
    Write-Output "Please reboot and sign in with the AAD account ($targetProfile). Then re-run this script with the -PostReboot flag."
}

<#
.SYNOPSIS
    Completes the profile migration process (post-reboot phase).

.DESCRIPTION
    Executes the second phase of profile migration after the system has been rebooted
    and the user has signed in with the AAD account. This includes restoring credentials,
    repairing OneDrive paths, and triggering MDM synchronization.

.EXAMPLE
    Complete-PostRebootTasks
    Completes the migration process with interactive prompt for the target profile name.

.NOTES
    This function should be run after rebooting and signing in with the new AAD account.
    It will prompt for the target profile name and restore the previously saved credentials.
#>
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
