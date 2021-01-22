# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

<#
# for testing
[System.Reflection.Assembly]::LoadFrom(".\modules\OauthHelper\NetCoreAssemblies\Microsoft.IdentityModel.Clients.ActiveDirectory.dll")
$variables = (Get-Content -Path "local.settings.json" | ConvertFrom-Json).Values
"REFRESH_TOKEN","TENANT_NAME","APPLICATION_ID","APPLICATION_SECRET","TEAMS_NAME","SP_SITE","SP_LIST","SP_RESOURCE_PRINCIPAL", "SYNC_MONTHS_PAST", "SYNC_MONTHS_FUTURE" | ForEach-Object {
    Set-Item -Path "env:$($_)" -Value $variables.$_
}
#>

function Invoke-Synchronization
{
    # All configuration values should be in environment variables
    $tenantName = $env:TENANT_NAME
    $clientId = $env:APPLICATION_ID
    $clientSecret = $env:APPLICATION_SECRET
    $base64Certificate = $env:CERTIFICATE
    $certificatePassword = $env:CERTIFICATE_PASSWORD
    $sharePointResourcePrincipal = $env:SP_RESOURCE_PRINCIPAL
    $site = $env:SP_SITE
    $listName = $env:SP_LIST
    $shiftsTeamName = $env:TEAMS_NAME
    $syncMonthsPast = [int]$env:SYNC_MONTHS_PAST
    $syncMonthsFuture = [int]$env:SYNC_MONTHS_FUTURE

    $functionStartedAt = Get-Date

    # validate time ranges
    if($syncMonthsPast -lt 0) {
        $syncMonthsPast = 0
    }

    if($syncMonthsFuture -lt 0) {
        $syncMonthsFuture = 0
    }

    # Get access keys (application identity)
    $graphToken = Invoke-ClientCredentialsFlow -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -Resource "https://graph.microsoft.com" | ConvertTo-AuthorizationHeaders
    #$sharePointToken = Invoke-ClientCredentialsFlow -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -Resource $sharePointResourcePrincipal | ConvertTo-AuthorizationHeaders

    #$certPassword = "" | ConvertTo-SecureString -AsPlainText -Force
    #$certificate = Get-PfxCertificate -FilePath "SharePointShiftsSync.pfx" -Password $certPassword
    if($certificatePassword -and $certificatePassword.Length.Trim() -gt 0) {
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $certificateSecurePassword = $certificatePassword | ConvertTo-SecureString -AsPlainText -Force
        $certificate.Import([System.Convert]::FromBase64String($base64Certificate), $certificateSecurePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]"DefaultKeySet")
    } else {
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]([System.Convert]::FromBase64String($base64Certificate))
    }

    $sharePointToken = Invoke-ClientCredentalsCertificateFlow -Tenant $tenantName -ClientId $clientId -Certificate $certificate -Resource $sharePointResourcePrincipal | ConvertTo-AuthorizationHeaders

    # Renew access keys (delegated auth)
    #$token = New-AccessToken -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -RefreshToken $refreshToken
    #$graphToken = Invoke-OnBehalfOfFlow -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -AccessToken $token.AccessToken -Resource "https://graph.microsoft.com" | ConvertTo-AuthorizationHeaders
    #$SharePointToken = Invoke-OnBehalfOfFlow -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -AccessToken $token.AccessToken -Resource $sharePointResourcePrincipal | ConvertTo-AuthorizationHeaders
    
    # Set time range
    $today = Get-Date
    $currentMonth = [DateTime]::new($today.Year, $today.Month, 1)

    $StartDate = $currentMonth.AddMonths($syncMonthsPast * -1)
    $EndDate = $currentMonth.AddMonths($syncMonthsFuture + 1) # +1 to always sync current month

    Write-Host "Synchronization range is from $($StartDate.ToString("d")) to $($EndDate.ToString("d"))."

    # And sync Shifts in Teams team to SharePoint list
    Sync-ShiftsToSharePoint -StartDate $StartDate -EndDate $EndDate `
                            -Site $Site -ListName $ListName `
                            -SharePointToken $SharePointToken -GraphToken $GraphToken `
                            -TeamName $shiftsTeamName `
                            -ShortenLastDay $true `
                            -Debug:$true

    $scriptDuration = (Get-Date) - $functionStartedAt
    $message = "Synchronization finished after $($scriptDuration.TotalSeconds) seconds."
    Write-Host $message

    $message
}
