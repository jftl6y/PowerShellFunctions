using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."
Import-Module AzureAD -UseWindowsPowerShell 

Connect-AzureAD -TenantDomain 'yourtenant.onmicrosoft.com' -Credential $credential -AzureEnvironmentName AzureUSGovernment

# Interact with query parameters or the body of the request.
$invitedUserDisplayName = $Request.Query.name
if (-not $invitedUserDisplayName){
    $invitedUserDisplayName = $Request.Body.Name
}
$invitedUserEmailAddress = $Request.Query.email
if (-not $invitedUserEmailAddress){
    $invitedUserEmailAddress= $Request.Body.Email
}
New-AzureADMSInvitation -InvitedUserDisplayName $invitedUserDisplayName -InvitedUserEmailAddress $invitedUserEmailAddress -InviteRedirectURL https://myapps.microsoft.com -SendInvitationMessage $true

$body = "Created user $invitedUserEmailAddress"

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
