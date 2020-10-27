param
(
    [Parameter(Mandatory=$true)]
    $Username,

    [Parameter(Mandatory=$true)]
    $Password,

    [Parameter(Mandatory=$true)]
    $Team,
        
    [Parameter(Mandatory=$true)]
    $Channel
        
)


if ($Username -ne "" -and $Password -ne "" -and $Teams -ne "" -and $Channel -ne "")
{
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    $Date = Get-Date -Format "MM-dd-yyyy-HHmm"

    #parameters that need to be changed after Azure Ad config
    $clientId = "<ID>"
    $tenantName = "m365x706011.onmicrosoft.com"
    $clientSecret = "<SECRET>"
    $resource = "https://graph.microsoft.com/"

    $ReqTokenBody = @{
        Grant_Type    = "Password"
        client_Id     = $clientID
        Client_Secret = $clientSecret
        Username      = $Username
        Password      = $Password
        Scope         = "https://graph.microsoft.com/.default"
    } 
    $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

    #Getting the user object
    $apiUrl = "https://graph.microsoft.com/v1.0/me/"
    $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
    $UsernameId = $Data.id

    #Getting all Groups
    $apiUrl = "https://graph.microsoft.com/beta/groups/"
    $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
    $Groups = ($Data | Select-Object Value).Value

    $TeamID = ($Groups | Where-Object {$_.displayname -eq "$($Team)"}).id

    $apiUrl = "https://graph.microsoft.com/v1.0/teams/$TeamID/Channels"
    $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get

    $ChannelID = (($Data | Select-Object Value).Value | Where-Object {$_.displayName -eq "$($Channel)"}).ID

    Write-Host "Exporting messages for the Team $TeamID in channel ($ChannelID)" 

    $apiUrl = "https://graph.microsoft.com/beta/groups/$TeamID/members"
    $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
    $UsersIDs = ($Data | Select-Object Value).Value.ID 

    # Join $UsernameID to $TeamID if it is not a member
    if ($UsersIDs -notcontains $UsernameID){
        $apiUrl = "https://graph.microsoft.com/beta/groups/$TeamID/members/`$ref"
        $body = @"
{
    "@odata.id": "https://graph.microsoft.com/beta/directoryObjects/f2637e45-0c9d-406d-b1b8-7149ffa70e8e"
}
"@
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Post -ContentType 'application/json' -Body $body
        Start-Sleep -Seconds 3
    } 
        
        
            
    #messages from channel

    $apiUrl = "https://graph.microsoft.com/beta/teams/$TeamID/channels/$ChannelID/messages"
    $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get 
    $Messages = ($Data | Select-Object Value).Value
    $mess = $Messages | Select-Object @{Name = 'DateTime'; Expression = {Get-Date -Date (($_).createdDateTime) -Format 'dd/MM/yyyy HH:mm'}}, @{Name = 'From'; Expression = {((($_).from).user).displayName}}, @{Name = 'Message'; Expression = {(($_).body).content -replace '<.*?>',''}} | Sort-Object DateTime



    $Header = @"
<style>
h1, h5, th { text-align: center; } 
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; } 
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; } 
td { font-size: 11px; padding: 5px 20px; color: #000; } 
tr { background: #b8d1f3; } 
tr:nth-child(even) { background: #dae5f4; } 
tr:nth-child(odd) { background: #b8d1f3; }
</style>
"@

    $body = "<body><b>Generated:</b> $(Get-Date) <br><br> <b>Team Name:</b> $($Team) <br> <b>Channel Name:</b> $($Channel) <br><br>"
    $body = $body + "</head>"

    $messhtml = $Mess | ConvertTo-Html -Body $body -Head $Header
    $Export = "$dir\TeamsHistory\$Team-$Channel"
    New-Item -ItemType Directory -Path $Export -ErrorAction Ignore
    $messhtml | Out-File $Export\$Team-$Channel-$Date.html

    # Remove $UsernameID if it should not be a member of $TeamID
    if ($UsersIDs -notcontains $UsernameID){
        $apiUrl = "https://graph.microsoft.com/beta/groups/$TeamID/members/$UsernameID/`$ref"

        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Delete
    }
 
}
