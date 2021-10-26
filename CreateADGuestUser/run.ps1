using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

$azureTenantId= $env:tenantID

$returnResult = "Success"
try {
    $uri = "https://login.microsoftonline.us/$azureTenantId/oauth2/v2.0/token"
    $body = @{
            client_Id = "$env:clientId"
            scope = "https://graph.microsoft.us/.default"
            client_secret = "$env:clientSecret"
            grant_type = "client_credentials"
        }

    $contentType = 'application/x-www-form-urlencoded' 
    $aadToken = Invoke-WebRequest -Uri $uri -Body $body -Method Post -ContentType $contentType
    $aadToken = (ConvertFrom-Json $aadToken.Content).access_token

    $invitedUserDisplayName = $Request.Query.name
    if (-not $invitedUserDisplayName){
        $invitedUserDisplayName = $Request.Body.Name
    }
    $invitedUserEmailAddress = $Request.Query.email
    if (-not $invitedUserEmailAddress){
        $invitedUserEmailAddress= $Request.Body.Email
    }
    
    $graphUri = "https://graph.microsoft.us/v1.0"
    $invitationSuffix = "/invitations"
    $invitationContentType = "application/json"
    $invitationUri = $graphUri + $invitationSuffix

    $body = @{
        invitedUserEmailAddress = $invitedUserEmailAddress
        sendInvitationMessage = $true
        invitedUserMessageInfo = @{
            messageLanguage = "en-US"
            customizedMessageBody = "$env:MessageBody"
        }
        inviteRedirectUrl = "https://myapps.microsoft.com?tenantId="
    }
    $body = ConvertTo-Json $body

   Invoke-WebRequest -Method Post -Uri $invitationUri -ContentType $invitationContentType -Body $body -Headers @{"Authorization" = "Bearer $($aadToken)"} 
    $returnCode = [HttpStatusCode]::OK
}
catch {
    $returnResult = $Error
    $returnCode = [HttpStatusCode]::BadRequest
}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $returnCode
    Body = $returnResult
})
