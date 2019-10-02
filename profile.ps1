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

function Invoke-Synchronization
{
    # All configuration values should be in environment variables
    $tenantName = $env:TENANT_NAME
    $clientId = $env:APPLICATION_ID
    $clientSecret = $env:APPLICATION_SECRET
    $refreshToken = $env:REFRESH_TOKEN
    $sharePointResourcePrincipal = $env:SP_RESOURCE_PRINCIPAL
    $site = $env:SP_SITE
    $listName = $env:SP_LIST
    $shiftsTeamName = $env:TEAMS_NAME

    $functionStartedAt = Get-Date


    # Generate access keys
    $token = New-AccessToken -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -RefreshToken $refreshToken
    $graphToken = New-OnBehalfOfAccessToken -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -AccessToken $token.AccessToken -ResourcePrincial "https://graph.microsoft.com"
    $SharePointToken = New-OnBehalfOfAccessToken -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -AccessToken $token.AccessToken -ResourcePrincial $sharePointResourcePrincipal

    # Set time range
    $today = Get-Date
    $today = $today.AddMonths(-2)
    $StartDate = [DateTime]::new($today.Year, $today.Month, 1)
    $EndDate = $StartDate.AddMonths(3)#.AddDays(-1)

    # And sync Shifts in Teams team to SharePoint list
    Sync-ShiftsToSharePoint -StartDate $StartDate -EndDate $EndDate `
                            -Site $Site -ListName $ListName `
                            -SharePointToken $SharePointToken -GraphToken $GraphToken `
                            -TeamName $shiftsTeamName `
                            -ShortenLastDay $true `
                            -Debug:$true

    #Wait-Debugger

    $scriptDuration = (Get-Date) - $functionStartedAt
    $message = "Synchronization finished after $($scriptDuration.TotalSeconds) seconds."
    Write-Host $message

    $message
}
