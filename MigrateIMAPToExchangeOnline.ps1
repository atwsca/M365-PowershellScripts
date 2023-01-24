# MigrateIMAPToExchangeOnline.ps1
# Version: 1.0
# Author: Dipak Parmar (dipak.tech)
# Description: A simple PowerShell script to migrate IMAP mailboxes to Exchange Online using the IMAP migration feature in Exchange Online.
# Disclaimer: This script is provided as-is with no warranty or support. Use at your own risk.

# This script requires the Exchange Online Management module to be installed. See https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps for more information.
# This script will require following information in CSV format:
# - The source IMAP server
# - The source IMAP port
# - The source IMAP Security (None, Ssl, Tls)
# - MaxConcurrentMigrations (Default is 50)
# - MaxConcurrentIncrementalSyncs (Default is 25)
# - Username
# - Password
# - Email address

# Example CSV file:
# EmailAddress,UserName,Password, IMAPServer, IMAPPort, IMAPSecurity, MaxConcurrentMigrations, MaxConcurrentIncrementalSyncs
# terrya@contoso.edu,terry.adams,1091990, imap.contoso.edu, 993, Tls, 50, 25
# annb@contoso.edu,ann.beebe,2111991, imap.contoso.edu, 993, Tls, 50, 25
# paulc@contoso.edu,paul.cannon,3281986, imap.contoso.edu, Tls, TLS, 50, 25

[CmdletBinding()]
param (
    # The CSV file containing the user information to migrate
    [Parameter(Mandatory = $true)]

    [ValidateScript({
            if (-Not ($_ | Test-Path -PathType Leaf) ) {
                throw "The path specified either does not exist or is not a file."
            }
            if ($_ -notmatch "(\.csv)") {
                throw "The file specified is not a CSV file."
            }
            return $true   
        })]
    $CSVFilePath,
    # The MaxConcurrentMigrations value to use for the migration (Will Override CSV value)
    [Parameter(Mandatory = $false)][int]$MaxConcurrentMigrations = 50,
    # The MaxConcurrentIncrementalSyncs value to use for the migration (Will Override CSV value)
    [Parameter(Mandatory = $false)][int]$MaxConcurrentIncrementalSyncs = 25,
    # The ReviewMigrationBeforeStarting value to use for the migration
    [Parameter(Mandatory = $false)][bool]$ReviewMigrationBeforeStarting = $true
)

# Function that offers to Install required modules
function Install-RequiredModules {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$RequiredModules
    )

    # Check if the required modules are installed
    $RequiredModulesInstalled = $true
    foreach ($RequiredModule in $RequiredModules) {
        if (!(Get-Module $RequiredModule -ListAvailable)) {
            $RequiredModulesInstalled = $false
        }
    }

    # If the required modules are not installed, offer to install them
    if ($RequiredModulesInstalled -eq $false) {
        Write-Host "The required modules are not installed. Do you want to install them now? Select M for manuall install instructions. (Y/N/M)" -ForegroundColor Yellow
        $InstallModules = Read-Host
        if ($InstallModules -eq "Y") {
            foreach ($RequiredModule in $RequiredModules) {
                Install-Module $RequiredModule -Force -Scope CurrentUser
                Write-Host "The required modules have been installed." -ForegroundColor Green
            }
        }
        elseif ($InstallModules -eq "M") {
            Write-Host "To install the required modules, run the following command in a PowerShell session:" -ForegroundColor Yellow
            Write-Host "Install-Module -Name $RequiredModules -Force -Scope CurrentUser`n" -ForegroundColor Yellow
            Write-Host "Once module is installed, run the follwing command to import module by running the following command:" -ForegroundColor Yellow
            Write-Host "Import-Module -Name $RequiredModules`n" -ForegroundColor Yellow
            Write-Host "If the module is already installed, you can typically skip this step and run Connect-ExchangeOnline without manually loading the module first." -ForegroundColor Yellow
            Write-Host "For more information, see https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps" -ForegroundColor Yellow
            exit
        }
        else {
            Write-Host "Please install the required modules and try again." -ForegroundColor Red
            exit
        }
    }
    
}

# Function that Validates existing connection to Exchange Online
function Confirm-ExchangeOnlineConnection {
    $ReturnObject = Get-ConnectionInformation | Select-Object State, TokenStatus, UserPrincipalName | Where-Object { $_.TokenStatus -eq "Active" -and $_.State -eq "Connected" }
    if ($ReturnObject) {
        Write-Host "Connected to Exchange Online with $($ReturnObject.UserPrincipalName)" -ForegroundColor Green
        return $ReturnObject
    }
}

# Function that connects to Exchange Online
function Connect-ToExchangeOnline {
    
 
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline
    # Connect-ExchangeOnline -ShowBanner:$false -DelegatedOrganization $OrganizationDomain // For future use
    $ExchangeOnlineConnection = Confirm-ExchangeOnlineConnection
    Write-Host "Connected to Exchange Online with $($ExchangeOnlineConnection.UserPrincipalName)" -ForegroundColor Green

}

# Function that validates the CSV file
function Confirm-CSVFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$CSVFile
    )

    # Check that the CSV file exists
    if (!(Test-Path $CSVFile)) {
        Write-Host "The CSV file $CSVFile does not exist. Please check the path and try again." -ForegroundColor Red
        exit
    }

    # Check that the CSV file is in the correct format
    $CSVFileContent = Import-Csv $CSVFile

    return $CSVFileContent | Select-Object -Property EmailAddress, UserName, Password, IMAPServer, IMAPPort, IMAPSecurity, MaxConcurrentMigrations, MaxConcurrentIncrementalSyncs -ErrorAction Stop


}

# Function that returns unique IMAP Configuration values from the CSV file. We need to create the MigrationEndpoint for each unique IMAP configuration
function Get-UniqueIMAPConfigurations {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$CSVFileContent
    )

    return $CSVFileContent | Select-Object -Property IMAPServer, IMAPPort, IMAPSecurity, MaxConcurrentMigrations, MaxConcurrentIncrementalSyncs | Sort-Object -Property IMAPServer, IMAPPort, IMAPSecurity, MaxConcurrentMigrations, MaxConcurrentIncrementalSyncs | Get-Unique -AsString 


}

# Function that Tests Migration Server Availability using Test-MigrationServerAvailability
function Test-MigrationServersAvailability {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$UniqueIMAPConfigurations
    )

    # StoreTestMigrationServersAvailability results in a hashtable
    $TestMigrationServersAvailabilityResults = @{}

    foreach ($IMAPConfiguration in $UniqueIMAPConfigurations) {
        Write-Host "Testing Migration Server Availability for $($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort) with $($IMAPConfiguration.IMAPSecurity)" -ForegroundColor Blue
        $TestResult = Test-MigrationServerAvailability -IMAP -RemoteServer $($IMAPConfiguration.IMAPServer) -Port $($IMAPConfiguration.IMAPPort) -Security $($IMAPConfiguration.IMAPSecurity)
        $TestMigrationServersAvailabilityResults.Add("$($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort)", $TestResult)
    }

    return $TestMigrationServersAvailabilityResults
}

# Function that creates the MigrationEndpoints using the IMAP Configuration values. It uses New-MigrationEndpoint to create the endpoints
function New-MigrationEndpoints {
    param (
        [Parameter(Mandatory = $true)]$UniqueIMAPConfigurations
    )

    foreach ($IMAPConfiguration in $UniqueIMAPConfigurations) {
        # Migration Endpoint Name is in the format of IMAPServer:IMAPPort:IMAPSecurity:MaxConcurrentMigrations:MaxConcurrentIncrementalSyncs
        # Check if the MaxConcurrentMigrations or MaxConcurrentIncrementalSyncs parameters are set. If so, use those values. Otherwise, use the values from the CSV file
        if ($MaxConcurrentMigrations -and $MaxConcurrentIncrementalSyncs) {
            Write-Host "Creating MigrationEndpoint for $($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort) using $($IMAPConfiguration.IMAPSecurity) with MaxConcurrentMigrations set to $MaxConcurrentMigrations and MaxConcurrentIncrementalSyncs set to $MaxConcurrentIncrementalSyncs" -ForegroundColor Blue
            New-MigrationEndpoint -Name "$($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort):$($IMAPConfiguration.IMAPSecurity):${MaxConcurrentMigrations}:$MaxConcurrentIncrementalSyncs" -IMAP -RemoteServer $($IMAPConfiguration.IMAPServer) -Port $($IMAPConfiguration.IMAPPort) -Security $($IMAPConfiguration.IMAPSecurity) -MaxConcurrentMigrations $MaxConcurrentMigrations -MaxConcurrentIncrementalSyncs $MaxConcurrentIncrementalSyncs -ErrorAction Stop
        }
        elseif ($MaxConcurrentMigrations) {
            Write-Host "Creating MigrationEndpoint for $($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort) using $($IMAPConfiguration.IMAPSecurity) with MaxConcurrentMigrations set to $MaxConcurrentMigrations" -ForegroundColor Blue
            New-MigrationEndpoint -Name "$($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort):$($IMAPConfiguration.IMAPSecurity):${MaxConcurrentMigrations}:$($IMAPConfiguration.MaxConcurrentIncrementalSyncs)" -IMAP -RemoteServer $($IMAPConfiguration.IMAPServer) -Port $($IMAPConfiguration.IMAPPort) -Security $($IMAPConfiguration.IMAPSecurity) -MaxConcurrentMigrations $MaxConcurrentMigrations -MaxConcurrentIncrementalSyncs $($IMAPConfiguration.MaxConcurrentIncrementalSyncs) -ErrorAction Stop
        }
        elseif ($MaxConcurrentIncrementalSyncs) {
            Write-Host "Creating MigrationEndpoint for $($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort) using $($IMAPConfiguration.IMAPSecurity) with MaxConcurrentIncrementalSyncs set to $MaxConcurrentIncrementalSyncs" -ForegroundColor Blue
            New-MigrationEndpoint -Name "$($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort):$($IMAPConfiguration.IMAPSecurity):$($IMAPConfiguration.MaxConcurrentMigrations):$MaxConcurrentIncrementalSyncs" -IMAP -RemoteServer $($IMAPConfiguration.IMAPServer) -Port $($IMAPConfiguration.IMAPPort) -Security $($IMAPConfiguration.IMAPSecurity) -MaxConcurrentMigrations $($IMAPConfiguration.MaxConcurrentMigrations) -MaxConcurrentIncrementalSyncs $MaxConcurrentIncrementalSyncs -ErrorAction Stop
        }
        else {
            Write-Host "Creating MigrationEndpoint for $($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort) using $($IMAPConfiguration.IMAPSecurity)" -ForegroundColor Blue
            New-MigrationEndpoint -Name "$($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort):$($IMAPConfiguration.IMAPSecurity):$($IMAPConfiguration.MaxConcurrentMigrations):$($IMAPConfiguration.MaxConcurrentIncrementalSyncs)" -IMAP -RemoteServer $($IMAPConfiguration.IMAPServer) -Port $($IMAPConfiguration.IMAPPort) -Security $($IMAPConfiguration.IMAPSecurity) -MaxConcurrentMigrations $($IMAPConfiguration.MaxConcurrentMigrations) -MaxConcurrentIncrementalSyncs $($IMAPConfiguration.MaxConcurrentIncrementalSyncs) -ErrorAction Stop
        }
    }
}

function Get-MigrationEndpoints {

    $MigrationEndpoints = Get-MigrationEndpoint -ErrorAction Stop

    Write-Host "Found $($MigrationEndpoints.Count) Existing MigrationEndpoints" -ForegroundColor Blue

    return $MigrationEndpoints
}

# Function that finds the new MigrationEndpoints that need to be created
function Find-NewMigrationEndpoints {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$UniqueIMAPConfigurations,
        [Parameter(Mandatory = $true)]$ExistingMigrationEndpoints	
    )
	
    $NewMigrationEndpoints = @()
    foreach ($IMAPConfiguration in $UniqueIMAPConfigurations) {
        $MigrationEndpointName = "$($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort):$($IMAPConfiguration.IMAPSecurity):$($IMAPConfiguration.MaxConcurrentMigrations):$($IMAPConfiguration.MaxConcurrentIncrementalSyncs)"
        if (-not ($ExistingMigrationEndpoints | Where-Object { $_.Identity -eq $MigrationEndpointName })) {
            $NewMigrationEndpoints += $IMAPConfiguration
        }
    }

    Write-Host "Found $($NewMigrationEndpoints.Count) New MigrationEndpoints" -ForegroundColor Blue
    return $NewMigrationEndpoints
}

# Function that returns existing MigrationBatches
function Get-MigrationBatches {
    $MigrationBatches = Get-MigrationBatch -ErrorAction Stop

    Write-Host "Found $($MigrationBatches.Count) Existing MigrationBatches" -ForegroundColor Blue

    return $MigrationBatches
}

# Function that creates migration batch
function New-MigrationBatches {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$CSVFileContent,
        [Parameter(Mandatory = $true)]$UniqueIMAPConfigurations,
        [Parameter(Mandatory = $true)]$ExistingMigrationBatches
    )

    # Counter to keep track of the number of batches created
    $BatchCounter = 1

    foreach ($IMAPConfiguration in $UniqueIMAPConfigurations) {
        # Get the MigrationEndpoint Name for the current IMAP Configuration
        $MigrationEndpointName = "$($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort):$($IMAPConfiguration.IMAPSecurity):$($IMAPConfiguration.MaxConcurrentMigrations):$($IMAPConfiguration.MaxConcurrentIncrementalSyncs)"
        # MigrationBatchName will consit of the MigrationEndpointName and the BatchCounter to make it unique, replace all ":" with "-"
        $MigrationEndpointNameCleaned = $MigrationEndpointName.Replace(":", "-")
        $MigrationBatchName = "$MigrationEndpointNameCleaned-$BatchCounter"

        # if the MigrationBatch already exists, skip it and increment the BatchCounter and skip to the next IMAP Configuration
        if ($ExistingMigrationBatches | Where-Object { $MigrationBatchName -eq $_.Identity }) {
            Write-Host "MigrationBatch $MigrationBatchName already exists, skipping creation." -ForegroundColor Blue
            $BatchCounter++
            continue
        }

        # Get the MigrationEndpoint object for the current IMAP Configuration
        $MigrationEndpoint = Get-MigrationEndpoint -Identity $MigrationEndpointName

        # Get the Mailboxes for the current IMAP Configuration
        $Mailboxes = $CSVFileContent | Where-Object { $_.IMAPServer -eq $($IMAPConfiguration.IMAPServer) -and $_.IMAPPort -eq $($IMAPConfiguration.IMAPPort) -and $_.IMAPSecurity -eq $($IMAPConfiguration.IMAPSecurity) -and $_.MaxConcurrentMigrations -eq $($IMAPConfiguration.MaxConcurrentMigrations) -and $_.MaxConcurrentIncrementalSyncs -eq $($IMAPConfiguration.MaxConcurrentIncrementalSyncs) } | Get-Unique
        
        # Create the MigrationBatch for the current IMAP Configuration
        Write-Host "Creating MigrationBatch $MigrationBatchName" -ForegroundColor Blue

        $Mailboxes | Format-Table -AutoSize

        # Only EmailAdress,Username,Password is required for IMAP Migration, remove other columns
        $Mailboxes = $Mailboxes | Select-Object -Property EmailAddress, Username, Password

        $CSVData = $Mailboxes | ConvertTo-Csv -NoTypeInformation -Delimiter "," | ForEach-Object { $_ -replace '"', '' } 

        # Also save the each batch csv to a file as csv for later use if needed then increment the BatchCounter
        $CSVData | Out-File -FilePath "$MigrationBatchName.csv"

        # type of $CSVData is required to be Byte[] // Fixme: This is not working
        # $CSVData = [System.Text.Encoding]::UTF8.GetBytes($CSVData)
        
        $BatchCounter++

        New-MigrationBatch -Name $MigrationBatchName -SourceEndpoint $MigrationEndpoint -CSVData ([System.IO.File]::ReadAllBytes("$MigrationBatchName.csv")) -ErrorAction Stop
    }
}

# Function that starts migration batch with option to review the migration batch before starting it
function Start-MigrationBatches {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$UniqueIMAPConfigurations,
        [Parameter(Mandatory = $false)]$ReviewMigrationBatch = $false
    )

    $BatchCounter = 1

    foreach ($IMAPConfiguration in $UniqueIMAPConfigurations) {
        # Get the MigrationEndpoint Name for the current IMAP Configuration
        $MigrationEndpointName = "$($IMAPConfiguration.IMAPServer):$($IMAPConfiguration.IMAPPort):$($IMAPConfiguration.IMAPSecurity):$($IMAPConfiguration.MaxConcurrentMigrations):$($IMAPConfiguration.MaxConcurrentIncrementalSyncs)"
        # MigrationBatchName will consit of the MigrationEndpointName and the BatchCounter to make it unique, replace all ":" with "-"
        $MigrationEndpointNameCleaned = $MigrationEndpointName.Replace(":", "-")
        $MigrationBatchName = "$MigrationEndpointNameCleaned-$BatchCounter"

        # Get the MigrationBatch object for the current IMAP Configuration
        $MigrationBatch = Get-MigrationBatch -Identity $MigrationBatchName

        # If the MigrationBatch is already started, skip it and increment the BatchCounter and skip to the next IMAP Configuration
        if ($MigrationBatch.Status -ne "Stopped") {
            Write-Host "MigrationBatch $MigrationBatchName is already started, skipping start." -ForegroundColor Blue
            $BatchCounter++
            continue
        }

        $MigrationBatch | Format-Table -AutoSize

        # Start the MigrationBatch for the current IMAP Configuration
        if ($ReviewMigrationBatch) {
            Write-Host "Reviewing MigrationBatch for $MigrationEndpointName" -ForegroundColor Yellow
            Get-MigrationStatistics | Format-Table -AutoSize
            # Prompt the user to start the migration batch
            $StartMigrationBatch = Read-Host "Do you want to start the migration batch for $MigrationEndpointName? (Y/N)"
            if ($StartMigrationBatch -eq "Y") {
                $MigrationBatch | Start-MigrationBatch -ErrorAction Stop
                $BatchCounter++
            }
            else {
                Write-Host "Skipping MigrationBatch for $MigrationEndpointName" -ForegroundColor Yellow
                $BatchCounter++
            }
        }
        else {
            Write-Host "Starting MigrationBatch for $MigrationEndpointName" -ForegroundColor Yellow
            $MigrationBatch | Start-MigrationBatch -ErrorAction Stop
            $BatchCounter++
        }
    }
}

# Helper function that shows the how to use the script
function Show-HowToUse {
    Write-Host "How to use this script:" -ForegroundColor Yellow
    Write-Host "1. Create a CSV file with the following columns:" -ForegroundColor Yellow
    Write-Host "   EmailAddress,Username,Password,IMAPServer,IMAPPort,IMAPSecurity,MaxConcurrentMigrations,MaxConcurrentIncrementalSyncs" -ForegroundColor Yellow
    Write-Host "2. Run the script with the following parameters:" -ForegroundColor Yellow
    Write-Host "   .\IMAPMigration.ps1 -CSVFilePath <Path to the CSV file> -ReviewMigrationBatch <True/False>" -ForegroundColor Yellow
}

# Use Install-RequiredModules function to install the required modules
Install-RequiredModules -RequiredModules @("ExchangeOnlineManagement")

# if ValidateConnection function returns false then connect to Exchange Online
if (-not (Confirm-ExchangeOnlineConnection)) {
    Connect-ToExchangeOnline -ExchangeOnlineConnection Confirm-ExchangeOnlineConnection
}

# Use Confirm-CSVFile function to validate the CSV file and return the CSV file content if it is valid
$CSVFileContent = Confirm-CSVFile -CSVFile $CSVFilePath

# if csvfilecontent is null then exit the script
if (-not $CSVFileContent) {
    Write-Host "Exiting because the CSV file is not valid" -ForegroundColor Red
    exit
}
# Use Get-UniqueIMAPConfigurations function to get the unique IMAP Configuration values from the CSV file
$UniqueIMAPConfigurations = Get-UniqueIMAPConfigurations -CSVFileContent $CSVFileContent

# Use Test-MigrationServersAvailability function to test the Migration Server Availability using the IMAP Configuration values
$TestResults = Test-MigrationServersAvailability -UniqueIMAPConfigurations $UniqueIMAPConfigurations

# TestResults is hashtable with key as IMAP Configuration and value as Test Result for the IMAP Configuration,  if the Test Result is not Success then exit the script
if ($TestResults) {
    foreach ($TestResult in $TestResults.GetEnumerator()) {
        if ("Success" -ne $TestResult.Value) {
            Write-Host "Test Result for $($TestResult.Key) is $($TestResult.Value)" -ForegroundColor Red
            # Write Exiting because the Test Result is not Success
            Write-Host "Exiting because the Test Result for $($TestResult.Key) is not Success" -ForegroundColor Red
            # Proivde Debug Steps what needs to be done to resolve the issue
            Write-Host "Please check the following steps to resolve the issue:" -ForegroundColor Yellow
            Write-Host "1. Check the IMAP Server is reachable from the Internet" -ForegroundColor Yellow
            Write-Host "2. Check the IMAP Server is reachable from the Exchange Online servers using the IMAP Port and IMAP Security" -ForegroundColor Yellow
            exit
        }
        else {
            Write-Host "Test Result for $($TestResult.Key) is $($TestResult.Value)" -ForegroundColor Green
        }
    }
}

# Use Get-MigrationEndpoints function to get the existing MigrationEndpoints
$ExistingMigrationEndpoints = Get-MigrationEndpoints

# Use Find-NewMigrationEndpoints function to get the new MigrationEndpoints using the IMAP Configuration values
$NewMigrationEndpoints = Find-NewMigrationEndpoints -UniqueIMAPConfigurations $UniqueIMAPConfigurations -ExistingMigrationEndpoints $ExistingMigrationEndpoints

# if NewMigrationEndpoints is not null then create the new MigrationEndpoints
if ($NewMigrationEndpoints) {
    # Use New-MigrationEndpoints function to create the new MigrationEndpoints
    New-MigrationEndpoints -UniqueIMAPConfigurations $NewMigrationEndpoints
}

# Use Find-ExistingMigrationBatches function to get the existing MigrationBatches using the IMAP Configuration values
$ExistingMigrationBatches = Get-MigrationBatches

$ExistingMigrationBatches | Format-Table -AutoSize

# Use New-MigrationBatches function to create the MigrationBatches using the IMAP Configuration values
$NewMigrationBatches = New-MigrationBatches -CSVFileContent $CSVFileContent -UniqueIMAPConfigurations $UniqueIMAPConfigurations -ExistingMigrationBatches $ExistingMigrationBatches

$NewMigrationBatches | Format-Table -AutoSize

# Use Start-MigrationBatches function to start the MigrationBatches using the IMAP Configuration values
Start-MigrationBatches -UniqueIMAPConfigurations $UniqueIMAPConfigurations -ReviewMigrationBatch $ReviewMigrationBeforeStarting