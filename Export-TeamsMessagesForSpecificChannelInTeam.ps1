# Export-TeamsMessagesForSpecificChannelInTeam.ps1
# Version: 1.0
# Author: Dipak Parmar
# Description: This script will export all messages from a specific channel in a specific team to a JSON file.
# Credits: https://pnp.github.io/cli-microsoft365/sample-scripts/teams/export-teams-conversations/
# Usage:
# 1. Install CLI for Microsoft 365: https://pnp.github.io/cli-microsoft365/
# 2. Login to your tenant using: m365 login
# 3. Run the script
# 4. Provide the Team ID and Channel ID when prompted

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true) ] [string] $teamId,
    [Parameter(Mandatory = $true) ] [string] $channelId
)


function  Get-Messages {
    param (
        [Parameter(Mandatory = $true)] [string] $teamId,
        [Parameter(Mandatory = $true)] [string] $channelId
    )
    $messages = m365 teams message list --teamId $teamId --channelId $channelId -o json | ConvertFrom-Json -AsHashtable
    return $messages
}
function  Get-MessageReplies {
    param (
        [Parameter(Mandatory = $true)] [string] $teamId,
        [Parameter(Mandatory = $true)] [string] $channelId,
        [Parameter(Mandatory = $true)] [string] $messageId
    )
  
    $messageReplies = m365 teams message reply list --teamId $teamId --channelId $channelId --messageId $messageId -o json | ConvertFrom-Json -AsHashtable
    return $messageReplies
}


$messages = Get-Messages $teamId $channelId
$messagesCollection = [System.Collections.ArrayList]@()
foreach ($message in $messages) {
    $messageReplies = Get-MessageReplies $teamId $channelId $message.id
    $messageDetails = $message
    [void]$messageDetails.Add("replies", $messageReplies)
    [void]$messagesCollection.Add($messageDetails)
}

$output = @{}
[void]$output.Add("messages", $messagesCollection)
$executionDir = $PSScriptRoot
$outputFilePath = "$executionDir/$(get-date -f yyyyMMdd-HHmmss).json"
# ConvertTo-Json cuts off data when exporting to JSON if it nests too deep. The default value of Depth parameter is 2. Set your -Depth parameter whatever depth you need to preserve your data.
$output | ConvertTo-Json -Depth 100 | Out-File $outputFilePath 
Write-Host "Open $outputFilePath file to review your output" -F Green 
