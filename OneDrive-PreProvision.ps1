$IsAzureADInstalled = false;
$IsMSOnlineInstalled = false;

If($IsWindows) {
    If( Get-InstalledModule AzureAD -ErrorAction SilentlyContinue) {
    $IsMicrosoftTeamsInstalled = true;
    Write-Host "Checking ... AzureAD Module is Installed`n"
}
else {
    try {
        Write-Host "Installing AzureAD Module`n"
        Install-Module -name AzureAD -Force -ErrorAction SilentlyContinue
        Write-Host "AzureAD Module is now Installed`n"
        $IsMicrosoftTeamsInstalled = true;
    }
    catch [Exception] {
        $_.message
        exit 
    }
}

If( Get-InstalledModule MSOnline -ErrorAction SilentlyContinue) {
    $IsMSOnlineInstalled = true;
    Write-Host "Checking ... MSOnline Module is Installed`n"
}
else {
    try {
        Write-Host "Installing MSOnline Module`n"
        Install-Module MSOnline -Force -ErrorAction SilentlyContinue
        Write-Host "MSOnline Module is now Installed`n"
        $IsMSOnlineInstalled = true;
    }
    catch [Exception] {
        $_.message
        exit 
    }
}

Write-Host "Please input admin account Credentials`n"

$Credential = Get-Credential
Connect-MsolService -Credential $Credential
$SharepointAdminurl = Read-Host -Prompt 'Input your sharepoint Admin URL with https (you can find this from your sharepoint admin center)'
Connect-SPOService -Credential $Credential -Url $SharepointAdminurl

$list = @()
#Counters
$i = 0


#Get licensed users
$users = Get-MsolUser -All | Where-Object { $_.islicensed -eq $true }
#total licensed users
$count = $users.count

foreach ($u in $users) {
    $i++
    Write-Host "$i/$count"

    $upn = $u.userprincipalname
    $list += $upn

    if ($i -eq 199) {
        #We reached the limit
        Request-SPOPersonalSite -UserEmails $list -NoWait
        Start-Sleep -Milliseconds 655
        $list = @()
        $i = 0
    }
}

if ($i -gt 0) {
    Request-SPOPersonalSite -UserEmails $list -NoWait
}

}
else {
    try {
        Write-Host "This Powershell script requires modules, which are only Windows compatible. Please use Windows machine to work with this script.`n"
    }
    catch [Exception] {
        $_.message
        exit 
    }
}

