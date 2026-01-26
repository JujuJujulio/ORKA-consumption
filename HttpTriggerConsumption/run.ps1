using namespace System.Net
#Import-Module ExchangeOnlineManagement #Mogelijke wijze
# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# define which code path to follow OK?
#   Exchange:   do the exchange code, cahinging aliases of mailbox
#   SayHello:   only give a response, returning the parameters
#   Version:    only return version and configuration info
$codePath = "Version"
switch ($codePath) {
    "Version" {
        $ps_version = $PSVersionTable.PSVersion -join "."
        $modules = Get-InstalledModule | ForEach-Object { "$($_.Name) ($($_.Version))" }
        $body = "PSVersion: $ps_version, Modules: $($modules -join ", ")"
    }
    "SayHello" { 
        # Interact with query parameters or the body of the request.
        $name = $Request.Query.Name
        if (-not $name) {
            $name = $Request.Body.Name
        }

        $body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
        if ($name) {
            $body = "Hello, $name. This HTTP triggered function executed successfully. CHaaaaaaaange Github"
        }
    }
    "Exchange" {
        $domain = "VBSDeKlimmuur.onmicrosoft.com"
        $upn = "rik.hendriks@$domain"
        $aliasesToAdd = "rh1@$domain", "rh2@$domain"
        $aliasesToRemove = "rh@$domain"
    
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
    }
    Default {
        $body = "Default codePath"
    }
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})

#Actuele code
#New push