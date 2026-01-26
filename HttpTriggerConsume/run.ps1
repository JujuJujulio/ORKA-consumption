using namespace System.Net
#Import-Module ExchangeOnlineManagement #Mogelijke wijze
# Input bindings are passed in via param block.
param($Request, $TriggerMetadata) #idk youtube vid met uitleg over param blokken in azure functions powershell

# Input parameters
$upn = $Request.Body.upn
$aliasesToAdd = $Request.Body.AliasesAdd | ForEach-Object { $_.Trim() }
$aliasesToRemove = $Request.Body.AliasesRm | ForEach-Object { $_.Trim() }

# Log parameters
Write-Host "Upn: $upn" 
Write-Host "AliasesToAdd: $($aliasesToAdd -join ", ")"
Write-Host "AliasesToRemove: $($aliasesToRemove -join ", ")"

# 1. Add and remove aliases as requested
Set-Mailbox $upn -EmailAddresses @{Add=$aliasesToAdd;Remove=$aliasesToRemove}    

# 2. Return actual situation
$mbox = get-mailbox -Identity $upn
if ($null -ne $mbox) {
    $aliases = $mbox.emailaddresses | Where-Object { $_ -clike "smtp:*"} | ForEach-Object { $_.substring(5) } 
    $message = "'$upn' aliases: $($aliases -join ", ")"
    $status = "Success"
} else {
    $message = "Error '$upn': mailbox does not exist. $($error[0].exception.message)"
    $status = "Error"
}    

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ 
    StatusCode = 200 
    Body       = @{  
        message = $message  
        status = $status
    } 
})
#From gitBash

#Call als volgt:
# Dooor ai => curl -X POST "http://localhost:7071/api/HttpTriggerConsume" -H "Content-Type: application/json" -d "{\"upn\":\"testuser\",\"AliasesAdd\":[\"alias1\",\"alias2\"],\"AliasesRm\":[\"oldalias1\"]}"

#$uri = "https://jouw-app.azurewebsites.net" 
 
# Definieer de parameters in een hashtable 
#$body = @{ 
#    upn           = "user@example.com" 
#    addAliases    = @("alias1@example.com", "alias2@example.com") 
#    removeAliases = @("oldalias@example.com") 
#} 
 
## Zet de hashtable om naar JSON en maak de call 
#$jsonBody = $body | ConvertTo-Json 
#$response = Invoke-RestMethod -Uri $uri -Method Post -Body $jsonBody -ContentType "application/json" 
 
## Toon het resultaat 
#$response.message