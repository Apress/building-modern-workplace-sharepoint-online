#Connect to Microsoft 365
Connect-PnPOnline -Scopes "Group.ReadWrite.All"

# Get access token for Graph
$accessToken = Get-PnPGraphAccessToken

#Pass access token to headers
$headers = @{
"Content-Type" = "application/json"
Authorization = "Bearer $accessToken"
} 

# Get list of M365 groups
$groupResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups" -Method Get -Headers $headers -UseBasicParsing

#filter Target group
$targetGroup = $groupResponse.value | Where-Object -FilterScript {$_.DisplayName -EQ 'team1'}

#get group id
$groupId = $targetGroup.id

#Create team for service executives.

$serviceTeam1 = @{
memberSettings = @{
allowCreateUpdateChannels = $true
}
messagingSettings = @{
allowUserEditMessages = $true
allowUserDeleteMessages = $true
}
funSettings = @{
allowGiphy = $true
giphyContentRating = "strict"
    allowStickersAndMemes= $true
    allowCustomMemes= $true
}
}
$serviceTeamBody1 = ConvertTo-Json -InputObject $serviceTeam1

$newTeam = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/team" -Method PUT -Headers $headers -Body $serviceTeamBody1 -UseBasicParsing

Write-Host $newTeam

