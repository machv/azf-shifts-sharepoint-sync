# Function App to sync teams to SharePoint
This PowerShell Function App synchronizes Shifts from Teams to SharePoint Calendar list.

## Deploy Function App

These PowerShell modules are required to be installed on the machine in order to deploy this app. All of them can be downloaded from the PowerShell Gallery.

- `AzureAD` - to create App Registration
- `Az` - to deploy Azure resources
- `OAuth2Toolkit` - to get OAuth2 tokens from Azure AD

```powershell
Install-Module AzureAD, Az, OAuth2Toolkit
```

### Required user accounts

| Reason                   | Required rights  |
| ------------------------ | ---------------- |
| Service account for sync | User that is Member of a Teams team and RW rights on SharePoint List |
| Register Azure AD app    | Standard user is enough to manage App Registration |
| Approve App registration | Global Admin |
| Deploy function app      | Contributor on a subscription |

### Azure AD Application Registration

To avoid conflicts in loading used DLL assemblies by used PowerShell modules (ADAL DLLs), we need to load PowerShell modules manually in expected order.

```powershell
Import-Module Az.Accounts
Import-Module AzureAD
Import-Module OAuth2Toolkit
```

This Function App requires to have an App Registration in the Azure AD tenant where a SharePoint Site and Teams team are.

```powershell
$appName = "Shifts to SharePoint Synchronization" # you can change the application name

# You can connect even as a standard user to register an Azure AD Application 
Connect-AzureAD
$session = Get-AzureADCurrentSessionInfo

$replyUrls = @("https://localhost:15484/auth")
$currentAadUser = Get-AzureADUser -ObjectId $session.Account

# Required Microsoft Graph permissions for the application
$resourceAccess = @(
    # used GUIDs are well-known for each resource. On way to obtain those is to see application definition using Graph Explorer (https://developer.microsoft.com/en-us/graph/graph-explorer)
    New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "Scope" # User.Read
    New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "5f8c59db-677d-491f-a6b8-5f174b11ec1d", "Scope" # Groups.Read.All
    New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "89fe6a52-be36-487e-b7d8-d061c450a026", "Scope" # Sites.ReadWrite.All
    New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "a154be20-db9c-4678-8ab7-66f6cc099a59", "Scope" # User.Read.All
)
$requiredResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
$requiredResourceAccess.ResourceAppId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
$requiredResourceAccess.ResourceAccess = $resourceAccess

$app = New-AzureADApplication -DisplayName $appName -ReplyUrls $replyUrls -AvailableToOtherTenants $false -RequiredResourceAccess $requiredResourceAccess

# Set current user as application owner
Add-AzureADApplicationOwner -ObjectId $app.ObjectId -RefObjectId $currentAadUser.ObjectId

# Generate application secret valid for 10 years
$startDate = Get-Date
$endDate = $startDate.AddYears(10)
$secretKey = New-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId -CustomKeyIdentifier "Azure Function App" -StartDate $startDate -EndDate $endDate

# And we need to obtain admin consent for the tenant for this application
# You need to provide Global Admin credentials for this
# Make sure that you will wait a few seconds before granting the consent (e. g. 30 secs) 
# to be sure AAD will be aware of the newly created application 
Invoke-AdminConsentForApplication -ClientId $app.AppId -RedirectUrl $replyUrls[0] -Tenant $session.TenantDomain
```

### Create a Function App

As the AAD Application is registered we can now deploy the Function App.

Change values in `$parameters` hash table to reflect your own environment (Teams team, SharePoint list etc.).

```powershell
$resourceGroupName = "litware-spsync-rg" # Set your destination resourge group
$location = "West Europe"

Connect-AzAccount
$currentContext = Get-AzContext
$currentUserGuid = (Get-AzADUser -UserPrincipalName $currentContext.Account).Id

$resourceGroup = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue
if(-not $ResourceGroup) 
{
    $resourceGroup = New-AzResourceGroup -Name $resourceGroupName -Location $location
}

$parameters = @{
    userGuid = $currentUserGuid
    resourceGroup = $resourceGroupName
    namePrefix = "litware-spsync"
    storageAccountName = "litwarespsyncdata"
    appSharePointListName = "Shifts-List"
    appSharePointResourcePrincipal = "https://<tenant>.sharepoint.com"
    appSharePointSiteUrl = "https://<tenant>.sharepoint.com/<site>/"
    appTeamsTeamName = "Contoso IT"
    appApplicationId = $app.AppId
    appTenantName = $session.TenantDomain
}
$deployment = New-AzResourceGroupDeployment -ResourceGroupName $resourceGroupName -Name $parameters["namePrefix"] -TemplateFile "./deployment/arm.json" -TemplateParameterObject $parameters -SkipTemplateParameterPrompt -Verbose
$functionAppName = $deployment.Outputs["appName"].Value
$keyVaultName = $deployment.Outputs["keyVaultName"].Value
```

### Obtain secrets and store them in Key Vault

In the login window login with service account credentials that is member of Teams team where shifts are stored and also don't forget to grant RW on SharePoint site with a calendar to that service account.

```powershell
# Get Refresh Token for the service account
$response = Invoke-CodeGrantFlow -RedirectUrl $replyUrls[0] -ClientId $app.AppId -ClientSecret $secretKey.Value -Tenant $session.TenantDomain -Resource $app.AppId -AlwaysPrompt $true

# Store secrets to Key Vault
$refreshTokenSecureString = ConvertTo-SecureString $response.refresh_token -AsPlainText -Force
$clientSecretSecureString = ConvertTo-SecureString $secretKey.Value -AsPlainText -Force

# Store them in a Key Vault
$refreshTokenSecret = Set-AzKeyVaultSecret -VaultName $keyVaultName -Name 'RefreshToken' -SecretValue $refreshTokenSecureString
$clientSecretSecret = Set-AzKeyVaultSecret -VaultName $keyVaultName -Name 'ClientSecret' -SecretValue $clientSecretSecureString

# And update Function App configuration to use those values from a Key Vault
$functionApp = Get-AzWebApp -ResourceGroupName $resourceGroupName -Name $functionAppName
$newAppSettings = @{}
foreach ($item in $functionApp.SiteConfig.AppSettings)
{
    $newAppSettings[$item.Name] = $item.Value
}
$newAppSettings["REFRESH_TOKEN"] = "@Microsoft.KeyVault(SecretUri=$($refreshTokenSecret.Id))"
$newAppSettings["APPLICATION_SECRET"] = "@Microsoft.KeyVault(SecretUri=$($clientSecretSecret.Id))"

Set-AzWebApp -ResourceGroupName $resourceGroupName -Name $functionAppName -AppSettings $newAppSettings
```

### Deploy code into the Function App

Finally we can deploy the code itself to Azure Function App.

```powershell
# Create ZIP file with app content
$zipContent = @(
    "HttpTrigger",
    "modules",
    "TimerTrigger",
    ".funcignore",
    "host.json",
    "profile.ps1",
    "proxies.json",
    "requirements.psd1"
)

$parent = [System.IO.Path]::GetTempPath()
$name = [System.IO.Path]::GetRandomFileName()
$tmp = New-Item -ItemType Directory -Path (Join-Path $parent $name)
$deploymentFile = Join-Path $tmp.FullName "Deployment.zip"

$zipContent | ForEach-Object {
    Copy-Item -Path $_ -Destination $tmp.FullName -Recurse
}

Compress-Archive -Path "$($tmp.FullName)/*" -DestinationPath $deploymentFile

# Publish
Publish-AzWebapp -WebApp $functionApp -ArchivePath $deploymentFile

# Cleanup
Remove-Item $tmp.FullName -Recurse
```

### Upgrade to newer version
To upgrade existing Function App to newer code just execute a section [Deploy code into the Function App](#deploy-code-into-the-function-app) again.

## For debugging

### Load data to JSON
```powershell
$variables = (Get-Content -Path "local.settings.json" | ConvertFrom-Json).Values

$today = (Get-Date).AddMonths(-2)
$startDate = [DateTime]::new($today.Year, $today.Month, 1)
$endDate = $StartDate.AddMonths(3)

$tenantName = $variables.TENANT_NAME
$clientId = $variables.APPLICATION_ID
$clientSecret = $variables.APPLICATION_SECRET
$refreshToken = $variables.REFRESH_TOKEN
$shiftsTeamName = $variables.TEAMS_NAME

$token = New-AccessToken -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -RefreshToken $refreshToken
$graphToken = Invoke-OnBehalfOfFlow -Tenant $tenantName -ClientId $clientId -ClientSecret $clientSecret -AccessToken $token.AccessToken -Resource "https://graph.microsoft.com" | ConvertTo-AuthorizationHeaders

$uri = "https://graph.microsoft.com/v1.0/me/joinedTeams"
$response = Invoke-RestMethod -Method Get -Headers $graphToken -Uri $uri
$team = $response.value | Where-Object { $_.displayName -eq $shiftsTeamName }

$ci = [CultureInfo]::InvariantCulture
$uri = "https://graph.microsoft.com/beta/teams/{0}/schedule/shifts?`$filter=sharedShift/startDateTime ge {1}T00:00:00.000Z and sharedShift/endDateTime le {2}T00:00:00.000Z" -f $team.id, $startDate.ToString("yyyy-MM-dd", $ci), $endDate.ToString("yyyy-MM-dd", $ci)
$response = Invoke-WebRequest -Method Get -Headers $graphToken -Uri $uri
Set-Content -Path "shifts.json" -Value $response.Content

$shifts = $response.Content | ConvertFrom-Json
foreach($shift in $shifts.value) {
    $uri = "https://graph.microsoft.com/v1.0/users/{0}" -f $shift.userId
    $userResponse = Invoke-WebRequest -Method Get -Headers $graphToken -Uri $uri
    Set-Content -Path "user_$($shift.userId).json" -Value $userResponse.Content 
}
```

### Import users from JSON
```powershell
Connect-AzureAD
$domain = (Get-AzureADDomain | where IsInitial -eq True).Name

$import = Get-ChildItem user_*.json
$import | foreach {
	$user = Get-Content $_ | ConvertFrom-Json

	$upn = $user.userPrincipalName.Replace("@domain.tld", "@" + $domain)
	
	$existing = Get-AzureADUser -Filter "userPrincipalName eq '$upn'"
	if(-not $existing) { 
		$passwordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
		$passwordProfile.Password = "FDSFDSFDSFSDFSD@4545"

		New-AzureADUser -DisplayName $user.displayName -GivenName $user.givenName -JobTitle $user.jobTitle -Mobile $user.mobilePhone -Surname $user.surname -UserPrincipalName $upn -AccountEnabled $true -MailNickName ($user.mail.Split("@")[0]) -PasswordProfile $passwordProfile
	}
}
```

### Mock functions for debugging
```powershell
function Get-Shifts
{
    $response = Get-Content "E:\Temp\Dummy Data\shifts.json" | ConvertFrom-Json    
    $response.value
}

function Get-UserDetails
{
    param(
        $Token,
        $UserId
    )

    $userResponse = Get-Content -Path "E:\Temp\Dummy data\user_$($UserId).json" | ConvertFrom-Json
    $userResponse.userPrincipalName = $userResponse.userPrincipalName.Replace("@domain.tld", "@demo.onmicrosoft.com")

    $userResponse
}
```
