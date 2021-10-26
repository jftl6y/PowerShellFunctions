using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

$azureTenantId= $env:tenantID
#$env:graphEndpoint is the Graph endpoint for the relevant cloud, e.g. graph.microsoft.com or graph.microsoft.us 
$graphUri = "$env:graphEndpoint/v1.0"
$graphScope = "$env:graphEndpoint/.default"



$returnResult = "Success"
try {
    #env:loginUri is the login endpoint for the relevant cloud, e.g. login.microsoftonline.com or login.microsoftonline.us
    $uri = "$env:loginUri/$azureTenantId/oauth2/v2.0/token"
    $body = @{
            client_Id = "$env:clientId"
            scope = "$graphScope"
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
        inviteRedirectUrl = "$env:RedirectUri"
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
