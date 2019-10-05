using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

Write-Host "Starting HTTP triggered Teams to SharePoint synchronization..."

#$environment = ls Env: | Out-String
#Write-Host $environment

try {
    $body = Invoke-Synchronization
    $status = [HttpStatusCode]::OK
}
catch {
    $status = [HttpStatusCode]::InternalServerError
    $body = $_.Message
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $status
    Body = $body
})
