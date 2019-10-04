# Function App to sync teams to SharePoint
This PowerShell Function App synchronizes Shifts from Teams to SharePoint Calendar list.

## Deploy Function App

To deploy this app these modules from PowerShell Gallery are needed to be on your computer installed:

- `AzureAD` - to create App Registration
- `Az` - to deploy Azure resources
- `OAuth2Toolkit` - to get OAuth2 tokens from Azure AD

### Azure AD Application Registration

To use this Function App we first need to have App Registration in Azure AD tenant where the SharePoint Site and Teams team is located.
 
```powershell
$appName = "Shifts to SharePoint Synchronization" # you can change the application name

Connect-AzureAD 
$session = Get-AzureADCurrentSessionInfo

$replyUrls = @("https://localhost:15484/auth")
$currentAadUser = Get-AzureADUser -ObjectId $session.Account

# Required Microsoft Graph permissions for the application
$resourceAccess = @(
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
Invoke-AdminConsentForApplication -ClientId $app.AppId -RedirectUrl $replyUrls[0] -Tenant $session.TenantDomain
```

### Deploy the Function App

As the AAD Application is registered we can now deploy the Function App.

Update corresponding variables and values in `$parameters` hash table to reflect your environment.

```powershell
$resourceGroupName = "function-app-rg" # Set your destination resourge group
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
    location = $Location
    resourceGroup = $resourceGroupName
    appName = "litware-spsync-app"
    keyVaultName = "litware-spsync-vault"
    storageAccountName = "litwarespsyncdata"
    appSharePointListName = "Ict-Info"
    appSharePointResourcePrincipal = "https://m365x074331.sharepoint.com"
    appSharePointSiteUrl = "https://m365x074331.sharepoint.com/it/"
    appTeamsTeamName = "Contoso IT"
    appApplicationId = $app.AppId
    appTenantName = $session.TenantDomain
}

$deployment = New-AzResourceGroupDeployment -ResourceGroupName $resourceGroupName -Name $parameters["appName"] -TemplateFile "./deployment/arm.json" -TemplateParameterObject $parameters -SkipTemplateParameterPrompt -Verbose
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
$refreshTokenSecret = Set-AzKeyVaultSecret -VaultName $parameters["keyVaultName"] -Name 'RefreshToken' -SecretValue $refreshTokenSecureString
$clientSecretSecret = Set-AzKeyVaultSecret -VaultName $parameters["keyVaultName"] -Name 'ClientSecret' -SecretValue $clientSecretSecureString

# And update Function App configuration to use those values from a Key Vault
$functionApp = Get-AzWebApp -ResourceGroupName $parameters["resourceGroup"] -Name $parameters["appName"]
$newAppSettings = @{}
foreach ($item in $functionApp.SiteConfig.AppSettings)
{
    $newAppSettings[$item.Name] = $item.Value
}
$newAppSettings["REFRESH_TOKEN"] = "@Microsoft.KeyVault(SecretUri=$($refreshTokenSecret.Id))"
$newAppSettings["APPLICATION_SECRET"] = "@Microsoft.KeyVault(SecretUri=$($clientSecretSecret.Id))"

Set-AzWebApp -ResourceGroupName $parameters["resourceGroup"] -Name $parameters["appName"]  -AppSettings $newAppSettings
```

### Deploy the Function App

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

## For debugging

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
