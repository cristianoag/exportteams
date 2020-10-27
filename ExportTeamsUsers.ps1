param
(
    [Parameter(Mandatory=$true)]
    $Username,
    [Parameter(Mandatory=$true)]
    $Password
)


if ($Username -ne "" -and $Password -ne "")
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
    #$UsernameID = "ecbb1541-e723-4265-9dd5-3f12d0e85d0b"

    Write-Host "Exporting chats for the user $Username ($UsernameID)" 

    #Getting all Chats for the user
    $apiUrl = "https://graph.microsoft.com/beta/users/$UsernameID/chats"
    $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
    $Chats = ($Data | Select-Object Value).Value



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

 
    $Export = "$dir\TeamsHistory\$Username"
    New-Item -ItemType Directory -Path $Export -ErrorAction Ignore

    $ChatIDs = ($Data | Select-Object Value).Value.ID 
    
    $count = 1
    foreach ($ChatID in $ChatIDs) {

        Write-Host "Exporting messages in chat $ChatID" 

        $body = "<body><b>Generated:</b> $(Get-Date) <br><br> <b>User Name:</b> $($Username) <br> <b>Chat ID:</b> $($ChatID) <br><br>"
        $body = $body + "</head>"

        $apiUrl = "https://graph.microsoft.com/beta/chats/$ChatID/messages"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get 
        $Messages = ($Data | Select-Object Value).Value

        $mess = $Messages | Select-Object @{Name = 'DateTime'; Expression = {Get-Date -Date (($_).createdDateTime) -Format 'dd/MM/yyyy HH:mm'}}, 
                                          @{Name = 'From'; Expression = {((($_).from).user).displayName}},
                                          @{Name = 'To'; Expression = {$Username}},
                                          @{Name = 'Message'; Expression = {(($_).body).content -replace '<.*?>',''}} | Sort-Object DateTime

        #$Messages 
        $messhtml = $mess | ConvertTo-Html -Body $body -Head $Header
        #$messhtml
        $messhtml | Out-File $Export\chat_$count.html
        $count++

    }
}