function New-AccessToken
{
    param(
        [string]$Tenant,
        [Parameter(ParameterSetName='ClientCredential')]
        [pscredential]$Client,
        [Parameter(ParameterSetName='ClientExplicit')]
        [string]$ClientId,
        [Parameter(ParameterSetName='ClientExplicit')]
        [string]$ClientSecret,
        [string]$RefreshToken
    )

    $authUrl = "https://login.microsoftonline.com/{0}/oauth2/token" -f $Tenant
    $parameters = @{
        grant_type = "refresh_token"
        client_secret= $ClientSecret
        refresh_token = $RefreshToken
        client_id = $ClientId
    }

    $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $parameters

    $expiresUtc = (Get-Date 01.01.1970)+([System.TimeSpan]::fromseconds($response.expires_on))
    $expires = [datetime]::SpecifyKind($expiresUtc, 'Utc').ToLocalTime()
    
    $result = [PSCustomObject]@{
        Expires = $expires
        AccessToken = $response.access_token
    }

    $result
}

<#
function Import-AdalLibrary
{
    $moduleName = "Az.Accounts" # AzureAD


    try 
    {
        $aadModule = Import-Module -Name $moduleName -ErrorAction Stop -PassThru
        $aadModule = Get-Module $moduleName # if already loaded previously by user, just select it
    }
    catch
    {
        throw "'Prerequisites not installed, Az.Accounts PowerShell module is required."
    }
    
    $dir = Join-Path $AadModule.ModuleBase "NetCoreAssemblies"
    $adal = Join-Path $dir "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
}
#>
function New-OnBehalfOfAccessToken 
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$Tenant,
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        [Parameter(Mandatory = $true)]
        [string]$clientSecret,
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,
        [Parameter()]
        [string]$ResourcePrincial = "https://graph.microsoft.com"
    )

    #Import-AdalLibrary

    $authority = "https://login.microsoftonline.com/$Tenant"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $clientCredentials = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList ($ClientId, $ClientSecret)
    $userAssertion  = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserAssertion" -ArgumentList ($AccessToken)
    $authResult = $authContext.AcquireTokenAsync($ResourcePrincial, $clientCredentials, $userAssertion)
    
    if ($authResult.Result.AccessToken) {
        # Creating header for Authorization token
        $authHeader = @{
            'Content-Type'  = 'application/json'
            'Authorization' = "Bearer " + $authResult.Result.AccessToken
            'ExpiresOn'     = $authResult.Result.ExpiresOn
        }
    
        $authHeader
    }
    elseif ($authResult.Exception) {
        throw "An error occured getting access token: $($authResult.Exception.InnerException)"
    }
}
