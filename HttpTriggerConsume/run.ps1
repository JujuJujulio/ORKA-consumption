using namespace System.Net
#Import-Module ExchangeOnlineManagement #Mogelijke wijze
# Input bindings are passed in via param block.
param($Request, $TriggerMetadata) #idk youtube vid met uitleg over param blokken in azure functions powershell

# Input parameters
$InputDomain = Request.Body.domain #Line 17 is hardcoded sometimes
$InputUpn = $Request.Body.upn
$InputAliasesToAdd = $Request.Body.AliasesAdd | ForEach-Object { $_.Trim() }
$InputAliasesToRemove = $Request.Body.AliasesRm | ForEach-Object { $_.Trim() }

# Log input parameters
Write-Host "InputUpn: $InputUpn"
Write-Host "InputAliasesToAdd: $($InputAliasesToAdd -join ", ")"
Write-Host "InputAliasesToRemove: $($InputAliasesToRemove -join ", ")"

# Organization connection settings
$domain = $InputDomain  # "VBSDeKlimmuur.onmicrosoft.com"
$upn = "$InputUpn@$domain"
$aliasesToAdd = $InputAliasesToAdd | ForEach-Object {"$_@$domain"}
$aliasesToRemove = $InputAliasesToRemove | ForEach-Object {"$_@$domain"}

# Log connection settings
Write-Host "Domain: $domain"
Write-Host "Upn: $upn" 
Write-Host "AliasesToAdd: $($aliasesToAdd -join ", ")"
Write-Host "AliasesToRemove: $($aliasesToRemove -join ", ")"

# 1. Add and remove aliases as requested
Set-Mailbox $upn -EmailAddresses @{Add=$aliasesToAdd;Remove=$aliasesToRemove}    

# 2. Return actual situation
$mbox = get-mailbox -Identity $upn
if ($null -ne $mbox) {
    $aliases = $mbox.emailaddresses | Where-Object { $_ -clike "smtp:*"} | ForEach-Object { $_.substring(5) } 
    $body = "'$upn' aliases: $($aliases -join ", ")"
} else {
    $body = "Error '$upn': mailbox does not exist. $($error[0].exception.message)"
}    


# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ 
    StatusCode = 200 
    Body       = @{  
        message = "Succesvol verwerkt voor $upn"  
        summary = "Toegevoegd: $($addAliases.Count), Verwijderd: $($removeAliases.Count)" 
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