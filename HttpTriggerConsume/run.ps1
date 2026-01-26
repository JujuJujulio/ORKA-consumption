using namespace System.Net
#Import-Module ExchangeOnlineManagement #Mogelijke wijze
# Input bindings are passed in via param block.
param($Request, $TriggerMetadata) #idk youtube vid met uitleg over param blokken in azure functions powershell

# Input parameters
#$InputDomain = Request.Query.domain #Line 17 is hardcoded for now
$InputUpn = $Request.Query.upn
if ($Request.Query.AliasesAdd){
    $InputAliasesToAdd = Request.Query.AliasesAdd -split "," | ForEach-Object { $_.Trim() }
}
if ($Request.Query.AliasesRm){
    $InputAliasesToRemove = Request.Query.AliasesToRm -split "," | ForEach-Object { $_.Trim() }
}
Write-Host "InputUpn: $InputUpn"
Write-Host "InputAliasesToAdd: $InputAliasesToAdd"
Write-Host "InputAliasesToRemove: $InputAliasesToRemove"

# Organization connection settings
$domain = "VBSDeKlimmuur.onmicrosoft.com"  # $InputDomain
$upn = "$InputUpn@$domain"
$aliasesToAdd = $InputAliasesToAdd | ForEach-Object {"$_$domain"}
$aliasesToRemove = $InputAliasesToRemove | ForEach-Object {"$_$domain"}

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
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
#From gitBash
