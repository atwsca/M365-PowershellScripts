[CmdletBinding()]
param (
    [Parameter(Mandatory = $True)]$Filename,
    [Parameter(Mandatory = $False)]$Location = "D:\HWID"
)

# Exit the script if os is not Windows
if ($IsWindows -eq $False) {
    Write-Error "This script only works on Windows"
    Exit
}

# Check if the filename has the .csv extension and if not add it to the filename
if ($Filename -notlike "*.csv") {
    $Filename = $Filename + ".csv"
}

# Check if the file exists in $Location Path and if so exit the script with an error
if (Test-Path -Path "$Location\$Filename") {
    Write-Error "File $Filename already exists in $Location"
    Exit
}

# Check if the directory exists and if not create it
if (!(Test-Path -Path $Location)) {
    Write-Host "Creating directory D:\HWID"
    New-Item -Type Directory -Path $Location
}

# If the location is not set to the directory, set it
if (!(Test-Path -Path $PWD.Path -Equals $Location)) {
    Write-Host "Changing directory to $Location"
    Set-Location -Path $Location
}

# Set the execution policy to allow the script to run
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned

# Install the Get-WindowsAutopilotInfo module
Install-Script -Name Get-WindowsAutopilotInfo

# Import the module
Import-Module -Name Get-WindowsAutopilotInfo

# Run the Get-WindowsAutopilotInfo command and save the output to the specified file
Get-WindowsAutopilotInfo -OutputFile $Filename
