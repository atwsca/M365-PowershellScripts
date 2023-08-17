# Get-AllTeamsMessagesExport.ps1
# Version: 1.0
# Author: Joseph Velliah (PnP Community)
# Description: This script will export all messages from all channels in all teams to a JSON file.
# Credits: https://pnp.github.io/cli-microsoft365/sample-scripts/teams/export-teams-conversations/
# Usage:
# 1. Install CLI for Microsoft 365: https://pnp.github.io/cli-microsoft365/
# 2. Login to your tenant using: m365 login
# 3. Run the script

function  Get-Teams {
  $teams = m365 teams team list -o json | ConvertFrom-Json -AsHashtable
  return $teams
}
function  Get-Channels {
  param (
    [Parameter(Mandatory = $true)] [string] $teamId
  )
  $channels = m365 teams channel list --teamId $teamId -o json | ConvertFrom-Json -AsHashtable
  return $channels
}
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

Try {
  $teamsCollection = [System.Collections.ArrayList]@()
  $teams = Get-Teams
  $progressCountTeam = 1;
  foreach ($team in $teams) {
    Write-Progress -Id 0 -Activity "Processing channels in Team : $($team.displayName)" -Status "Team $progressCountTeam of $($teams.length)" -PercentComplete (($progressCountTeam / $teams.length) * 100)
    $channelsCollection = [System.Collections.ArrayList]@()
    $channels = Get-Channels $team.id
    $progressCountChannel = 1;
    foreach ($channel in $channels) {
      Write-Progress -Id 1 -ParentId 0 -Activity "Processing messages in channel : $($channel.displayName)" -Status "Channel $progressCountChannel of $($channels.length)" -PercentComplete (($progressCountChannel / $channels.length) * 100)
      $messages = Get-Messages $team.id $channel.id
      $messagesCollection = [System.Collections.ArrayList]@()
      foreach ($message in $messages) {
        $messageReplies = Get-MessageReplies $team.id $channel.id $message.id
        $messageDetails = $message
        [void]$messageDetails.Add("replies", $messageReplies)
        [void]$messagesCollection.Add($messageDetails)
      }
      $channelDetails = $channel
      [void]$channelDetails.Add("messages", $messagesCollection)
      [void]$channelsCollection.Add($channelDetails)
      $progressCountChannel++
    }
    $teamDetails = $team
    [void]$teamDetails.Add("channels", $channelsCollection)
    [void]$teamsCollection.Add($teamDetails)
    $progressCountTeam++
  }
  Write-Progress -Id 0 -Activity " " -Status " " -Completed
  Write-Progress -Id 1 -Activity " " -Status " " -Completed
  $output = @{}
  [void]$output.Add("teams", $teamsCollection)
  $executionDir = $PSScriptRoot
  $outputFilePath = "$executionDir/$(get-date -f yyyyMMdd-HHmmss).json"
  # ConvertTo-Json cuts off data when exporting to JSON if it nests too deep. The default value of Depth parameter is 2. Set your -Depth parameter whatever depth you need to preserve your data.
  $output | ConvertTo-Json -Depth 100 | Out-File $outputFilePath 
  Write-Host "Open $outputFilePath file to review your output" -F Green 
}
Catch {
  $ErrorMessage = $_.Exception.Message
  Write-Error $ErrorMessage
}
