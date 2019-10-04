# Sharepoint helper functions
function Get-ItemTypeForListName($listName) 
{
    # Get List Item Type metadata
    $name = $listName.Replace("-", "")
    "SP.Data." + $name.SubString(0, 1).ToUpper() + ($name.Substring(1) -split " " -join "") + "ListItem"
}

function Invoke-EnsureSpUser
{
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]
        $User,
        [Parameter(Mandatory = $true)]
        $Site,
        [Parameter(Mandatory = $true)]
        $Token
    )

    $payload = [PSCustomObject]@{
        "logonName" = $User.userPrincipalName
    }
    $uri = "{0}/_api/web/ensureuser" -f $Site
    $response = Invoke-RestMethod -Method Post -Headers $Token -Uri $uri -Body ($payload | ConvertTo-Json)

    $response.d
}

function Add-SPShift
{
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "UserId")]
        $UserId,
        [Parameter(Mandatory = $true, ParameterSetName = "LookupUser")]
        $User,
        [Parameter(Mandatory = $true)]
        $Site,
        [Parameter(Mandatory = $true)]
        $Token,
        [Parameter(Mandatory = $true)]
        $ListName,
        [Parameter(Mandatory = $true)]
        [DateTime]$StartDate,
        [Parameter(Mandatory = $true)]
        [DateTime]$EndDate,
        [bool]$ShortenLastDay,
        [bool]$AsAllDayEvent = $true
    )

    if($PsCmdlet.ParameterSetName -eq "LookupUser")
    {
        # Ensure user in SharePoint site
        $spUser = Invoke-EnsureSpUser -User $User -Site $Site -Token $Token
        $spUserId = $spUser.Id
    }
    else {
        $spUserId = $UserId
    }

    # Insert item to the SharePoint list
    $itemType = Get-ItemTypeForListName -listName $ListName

    $start = $StartDate
    $end = $EndDate
    if($ShortenLastDay -and ([DateTime]$StartDate).Date -ne ([DateTime]$EndDate).Date) {
        $end = $EndDate.AddHours(-1 * $end.Hour).AddMinutes(-1 * $end.Minute).AddMinutes(-1)
        $end = $end.Date
    }

    if($AsAllDayEvent)
    {
        $start = $StartDate.Date
        $end = $end.Date
    }

    $payload = [PSCustomObject]@{
        "__metadata" = [PSCustomObject]@{ 
            "type" = $itemType
        }
        "EmployeeId" = $spUserId
        "Title" = ("{0} - {1}" -f $User.displayName, (Format-PhoneNumber $User.mobilePhone))
        "Description" = ("{0} - {1}" -f $StartDate, $EndDate)
        "EventDate" = $start.ToString("o")
        "EndDate" = $end.ToString("o") #"2019-08-20T08:32:51.451Z"
        "fAllDayEvent" = $AsAllDayEvent
    }
    # Due to bug in SharePoint API, one-day events can't be currently created as all day events (https://github.com/SharePoint/sp-dev-docs/issues/2755)
    if(([DateTime]$StartDate).Date -eq ([DateTime]$end).Date) {
         #$payload.fAllDayEvent = $false
    }

    $payload = ($payload | ConvertTo-Json)
    $payload = [System.Text.Encoding]::UTF8.GetBytes($payload)
    $uri = "{0}/_api/web/lists/GetByTitle('{1}')/items" -f $Site, $ListName
    $r = Invoke-WebRequest -Method Post -Headers $Token -Uri $uri -Body $payload

    Write-Debug (" * Inserting shift for {0} ({1} - {2})" -f $User.displayName, $StartDate, $EndDate)

    if($r.StatusCode -ne 201)
    {
        throw "Inserting of the item failed"
    }

    $response = ($r.Content | ConvertFrom-Json -AsHashtable).d
<#
    $item = Get-SpListItem -Id $response["ID"] -Site $Site -Token $SharePointToken -ListName $ListName 

    if(([DateTime]$item.EndDate) -ne $end)
    {
        $diff = $item.EndDate - $end

        $payload = [PSCustomObject]@{
            "__metadata" = [PSCustomObject]@{ 
                "type" = $itemType
            }
            "EndDate" = $end.ToString("o")
        }
        $payload = ($payload | ConvertTo-Json)
        $payload = [System.Text.Encoding]::UTF8.GetBytes($payload)

        $headers = New-ClonedObject -DeepCopyObject $SharePointToken
        $headers['X-HTTP-Method'] = "MERGE" 
        $headers['IF-MATCH'] = "*"

        $uri = "{0}/_api/web/lists/GetByTitle('{1}')/items({2})" -f $Site, $ListName, $response["ID"]
        $response = Invoke-WebRequest -Method Post -Headers $headers -Uri $uri -Body $payload
    
    }
    #>
    
    $response["ID"] # return newly created item id
}

function Get-SpListItem
{
    param(
        $Token,
        [string]$Site,
        [string]$ListName,
        $Id
    )

    $uri = "{0}/_api/web/lists/GetByTitle('{1}')/items({2})" -f $Site, $ListName, $Id
    $response = Invoke-WebRequest -Method Get -Headers $Token -Uri $uri

    $fixup = $response.Content -creplace '"Id":','"Id2":'
    $response = $fixup | ConvertFrom-Json

    $response.d
}

function Get-SpListItems
{
    [cmdletbinding()]
    param(
        $Token,
        [string]$Site,
        [string]$ListName,
        [DateTime]$StartDate,
        [DateTime]$EndDate
    ) 

    # https://www.c-sharpcorner.com/article/sharepoint-2013-using-rest-api-selecting-filtering-sortin/
    $uri = "{0}/_api/web/lists/GetByTitle('{1}')/items?`$filter=EventDate ge datetime'{2}' and EndDate le datetime'{3}'" -f $Site, $ListName, $StartDate.ToString("s", $ci), $EndDate.ToString("s", $ci)
    $response = Invoke-WebRequest -Method Get -Headers $Token -Uri $uri
    $fixup = $response.Content -creplace '"Id":','"Id2":'
    $response = $fixup | ConvertFrom-Json

    $response.d.results
}

function Remove-SpListItem
{
    [cmdletbinding()]
    param(
        $Token,
        [string]$Site,
        [string]$ListName,
        $ItemId
    )

    # Delete header set
    $headers = New-ClonedObject -DeepCopyObject $Token
    $headers['X-HTTP-Method'] = "DELETE" 
    $headers['IF-MATCH'] = "*"

    $uri = "{0}/_api/web/lists/GetByTitle('{1}')/items({2})" -f $Site, $ListName, $itemId
    $r = Invoke-RestMethod -Method Post -Headers $headers -Uri $uri
    
    $r
}

function Remove-SpListItems
{
    [cmdletbinding()]
    param(
        $Token,
        [string]$Site,
        [string]$ListName,
        [DateTime]$StartDate,
        [DateTime]$EndDate
    )

    $items = Get-SpListItems -Token $Token -Site $Site -ListName $ListName -StartDate $StartDate -EndDate $EndDate

    $item = $items | Select-Object -First 1
    foreach($item in $items) {
        $result = Remove-SpListItem -Token $Token -Site $Site -ListName $ListName -ItemId $item.ID
        Write-Debug (" * Removing item {0} ({1} - {2})" -f $item.Title, [DateTime]$item.EventDate, [DateTime]$item.EndDate)
    }
}

function New-ClonedObject 
{
    param($DeepCopyObject)
    $memStream = new-object IO.MemoryStream
    $formatter = new-object Runtime.Serialization.Formatters.Binary.BinaryFormatter
    $formatter.Serialize($memStream,$DeepCopyObject)
    $memStream.Position=0
    $formatter.Deserialize($memStream)
}

function Format-PhoneNumber
{
    param(
        [Parameter(Position=0)]
        [string]$Number
    )

    if($Number.Replace(" ", "") -match "^(\+420)?([0-9]{3})([0-9]{3})([0-9]{3})$") 
    {
        "{0} {1} {2}" -f $Matches[2], $Matches[3], $Matches[4]
    }
}

function Get-UserDetails
{
    param(
        $Token,
        $UserId
    )

    $uri = "https://graph.microsoft.com/v1.0/users/{0}" -f $UserId
    $userResponse = Invoke-RestMethod -Method Get -Headers $Token -Uri $uri

    $userResponse
}

function Get-Shifts
{
    param(
        [Parameter(Mandatory = $true)]
        $Token,
        [Parameter(Mandatory = $true)]
        $TeamId,
        [Parameter(Mandatory = $true)]
        [DateTime]$StartDate,
        [Parameter(Mandatory = $true)]
        [DateTime]$EndDate
    )
    $ci = [CultureInfo]::InvariantCulture
    $uri = "https://graph.microsoft.com/beta/teams/{0}/schedule/shifts?`$filter=sharedShift/startDateTime ge {1}T00:00:00.000Z and sharedShift/endDateTime le {2}T00:00:00.000Z" -f $TeamId, $StartDate.ToString("yyyy-MM-dd", $ci), $EndDate.ToString("yyyy-MM-dd", $ci)
    $response = Invoke-RestMethod -Method Get -Headers $Token -Uri $uri
    
    $response.value
}

function Get-TeamsTeam
{
    param(
        $Token,
        $Name
    )

    $uri = "https://graph.microsoft.com/v1.0/me/joinedTeams"
    $response = Invoke-RestMethod -Method Get -Headers $Token -Uri $uri
    $team = $response.value | Where-Object { $_.displayName -eq $Name } # Out-GridView -Title "Select correct team" -OutputMode Single

    if(-not $team) 
    {
        throw "Requested Team not found -> please check the name"
    }

    $team
}

function Add-ShiftsToSpList
{
    [cmdletbinding()]
    param(
        $Token,
        $Site,
        $ListName,
        $Shifts
    )

    $addOptions = @{
        SpSite = $Site
        SpToken = $Token
        SpListName = $ListName
    }
    
    $shortenLastDay = $true # if true shift will be ending by midnight of previous day (so following one will be only one there)
    $sortedShifts = $shifts <##| Where-Object {$_.userId -eq "79dfeb76-d715-43de-8c15-64df5d77ef24" }<##> | Sort-Object -Property userId, @{Expression = {$_.sharedShift.startDateTime } }
    $currentUserId = $null
    $pendingShiftStart = $null
    $pendingShiftEnd = $null
    $segments = @()
    $shift = $sortedShifts | Select-Object -Last 1
    foreach($shift in $sortedShifts) 
    {
        if($currentUserId -ne $shift.userId) 
        {
            if($pendingShiftStart -is [datetime]) 
            {
                # write
                $user = Get-UserDetails -Token $graphToken -UserId $currentUserId
                $r = Add-SPShift -ShortenLastDay:$shortenLastDay -User $user -StartDate $pendingShiftStart -EndDate $pendingShiftEnd @addOptions
                $segments | ForEach-Object { Write-Debug ("  > {0} - {1}" -f [DateTime]$_.sharedShift.startDateTime, [DateTime]$_.sharedShift.endDateTime) }
            }
    
            $currentUserId = $shift.userId
            $pendingShiftStart = [DateTime]$shift.sharedShift.startDateTime
            $pendingShiftEnd = [DateTime]$shift.sharedShift.endDateTime
            $segments = @($shift)
    
            continue
        }
    
        # if Date part is the same, then merge two together
        if($pendingShiftEnd.Date -eq ([DateTime]$shift.sharedShift.startDateTime).Date) 
        {
            $pendingShiftEnd = [DateTime]$shift.sharedShift.endDateTime
            $segments += $shift
        } 
        else 
        { # or create a separate and start a new one
            $user = Get-UserDetails -Token $graphToken -UserId $currentUserId
            $r = Add-SPShift -ShortenLastDay:$shortenLastDay -User $user -StartDate $pendingShiftStart -EndDate $pendingShiftEnd @addOptions
            $segments | ForEach-Object { Write-Debug ("  > {0} - {1}" -f [DateTime]$_.sharedShift.startDateTime, [DateTime]$_.sharedShift.endDateTime) }
    
            $pendingShiftStart = [DateTime]$shift.sharedShift.startDateTime
            $pendingShiftEnd = [DateTime]$shift.sharedShift.endDateTime  
            $segments = @($shift)
        }
    }
    if($pendingShiftStart) 
    {  
        $user = Get-UserDetails -Token $graphToken -UserId $currentUserId
        $r = Add-SPShift -ShortenLastDay:$shortenLastDay -User $user -StartDate $pendingShiftStart -EndDate $pendingShiftEnd @addOptions
        $segments | ForEach-Object { Write-Debug ("  > {0} - {1}" -f [DateTime]$_.sharedShift.startDateTime, [DateTime]$_.sharedShift.endDateTime) }
    }
}

function Complete-ShiftEvent
{
    param(
        [Parameter(Mandatory = $true)]
        $GraphToken,
        [Parameter(Mandatory = $true)]
        $SharePointToken,
        [Parameter(Mandatory = $true)]
        [string]$Site,
        [Parameter(Mandatory = $true)]
        [string]$ListName,
        [Parameter(Mandatory = $true)]
        $UserId,
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        [Parameter(Mandatory = $true)]
        [datetime]$EndDate,
        [Parameter(Mandatory = $true)]
        [bool]$ShortenLastDay,
        $Segments,
        $ExistingItems

    )

    $user = Get-UserDetails -Token $GraphToken -UserId $UserId
    $spUser = Invoke-EnsureSpUser -User $user -Site $Site -Token $SharePointToken

    $spItem = $ExistingItems | Where-Object { $_.EmployeeId -eq $spUser.Id -and $_.EventDate.Date -eq $StartDate.Date -and $_.EndDate.Date -eq $EndDate.Date -and $_.Description -eq ("{0} - {1}" -f $StartDate, $EndDate) }
    if($spItem -and $spItem -isnot [array])
    {
        Write-Debug ("Reusing existing item #{0} in the list as user ({2}) and duration ({1}) is same." -f $spItem.ID, $spItem.Description, $user.displayName)
        
        $itemId = $spItem.Id
    }
    else 
    {
        $itemId = Add-SPShift -ShortenLastDay $ShortenLastDay -User $user -StartDate $StartDate -EndDate $EndDate -Site $Site -Token $SharePointToken -ListName $ListName
        $segments | ForEach-Object { Write-Debug ("  > {0} - {1}" -f [DateTime]$_.sharedShift.startDateTime, [DateTime]$_.sharedShift.endDateTime) }
    }

    $itemId
}

function Sync-ShiftsToSharePoint
{
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]
        $SharePointToken,
        [Parameter(Mandatory = $true)]
        $GraphToken,
        [Parameter(Mandatory = $true)]
        [string]$TeamName,
        [Parameter(Mandatory = $true)]
        [string]$Site,
        [Parameter(Mandatory = $true)]
        [string]$ListName,
        [Parameter(Mandatory = $true)]
        [DateTime]$StartDate,
        [Parameter(Mandatory = $true)]
        [DateTime]$EndDate,
        [bool]$ShortenLastDay = $true,
        [bool]$SameDatePartMergingAllowed = $true,
        [switch]$ForceCleanup = $false
    )
<#
.SYNOPSIS

Synchornizes shifts from Teams team to SharePoint Calendar list for a time range

.EXAMPLE

PS> $token = New-AccessToken -Tenant $TenantName -ClientId $clientId -ClientSecret $clientSecret -RefreshToken $refreshToken
PS> $graphToken = New-OnBehalfOfAccessToken -Tenant $TenantName -ClientId $clientId -ClientSecret $clientSecret -AccessToken $token.AccessToken -ResourcePrincial "https://graph.microsoft.com"
PS> $SharePointToken = New-OnBehalfOfAccessToken -Tenant $TenantName -ClientId $clientId -ClientSecret $clientSecret -AccessToken $token.AccessToken -ResourcePrincial "https://<mytenant>.sharepoint.com"
PS> $today = Get-Date
PS> $StartDate = [DateTime]::new($today.Year, $today.Month, 1)
PS> $EndDate = $StartDate.AddMonths(3)
PS> 
PS> Sync-ShiftsToSharePoint -StartDate $StartDate -EndDate $EndDate `
PS>                         -Site "https://mytenant.sharepoint.com/site" -ListName "Events" `
PS>                         -SharePointToken $SharePointToken -GraphToken $GraphToken `
PS>                         -TeamName "Shifts Team" `
PS>                         -ShortenLastDay $true `
PS>                         -Debug:$true # with verbose logging output

#>

    $SharePointToken['Content-Type'] = "application/json;odata=verbose;charset=utf-8"
    $SharePointToken['Accept'] = "application/json;odata=verbose;charset=utf-8"
    
    $commonParameters = @{
        Token = $SharePointToken
        Site = $Site
        ListName = $ListName
    }

    $team = Get-TeamsTeam -Token $GraphToken -Name $TeamName
    $shifts = Get-Shifts -Token $GraphToken -TeamId $team.id -StartDate $StartDate -EndDate $EndDate
    $listItems = Get-SpListItems @commonParameters -StartDate $StartDate -EndDate $EndDate

    if($ForceCleanup) 
    {
        foreach($item in $listItems)
        {
            Write-Debug (" * Forcing removal of item {0} ({1} - {2})" -f $item.Title, [DateTime]$item.EventDate, [DateTime]$item.EndDate)
            $result = Remove-SpListItem @commonParameters -ItemId $item.ID
        }

        $listItems = @()
    }

    foreach($item in $listItems)
    {
        $item.Description = [System.Net.WebUtility]::HtmlDecode($item.Description)
    }

    # Main Logic
    $existingItems = @()
    $sortedShifts = $shifts <##| Where-Object {$_.userId -eq "79dfeb76-d715-43de-8c15-64df5d77ef24" }<##> | Sort-Object -Property userId, @{Expression = {$_.sharedShift.startDateTime } }
    $currentUserId = $null
    $pendingShiftStart = $null
    $pendingShiftEnd = $null
    $segments = @()
    $shift = $sortedShifts | Select-Object -First 1 -Skip 1
    #$DebugPreference = 'Continue'
    $completeParameters = @{
        GraphToken = $GraphToken
        SharePointToken = $SharePointToken
        Site = $Site
        ListName = $ListName
        ExistingItems = $listItems 
        ShortenLastDay = $shortenLastDay
    }
    foreach($shift in $sortedShifts) 
    {
        if($currentUserId -ne $shift.userId) 
        {
            if($pendingShiftStart -is [datetime]) 
            {
                # write
                $spItemId = Complete-ShiftEvent @completeParameters -UserId $currentUserId -StartDate $pendingShiftStart -EndDate $pendingShiftEnd -Segments $segments
                $existingItems += $spItemId
            }
    
            $currentUserId = $shift.userId
            $pendingShiftStart = [DateTime]$shift.sharedShift.startDateTime
            $pendingShiftEnd = [DateTime]$shift.sharedShift.endDateTime
            $segments = @($shift)
    
            continue
        }
      
        if($SameDatePartMergingAllowed -and $pendingShiftEnd.Date -eq ([DateTime]$shift.sharedShift.startDateTime).Date) 
        {
            # if Date part is the same, then merge two together
            $pendingShiftEnd = [DateTime]$shift.sharedShift.endDateTime
            $segments += $shift
        } 
        else 
        { 
            $spItemId = Complete-ShiftEvent @completeParameters -UserId $currentUserId -StartDate $pendingShiftStart -EndDate $pendingShiftEnd -Segments $segments
            $existingItems += $spItemId

            # start new one
            $pendingShiftStart = [DateTime]$shift.sharedShift.startDateTime
            $pendingShiftEnd = [DateTime]$shift.sharedShift.endDateTime  
            $segments = @($shift)
        }
    }

    # If there any leftovers then save that also
    if($pendingShiftStart) 
    {  
        $spItemId = Complete-ShiftEvent @completeParameters -UserId $currentUserId -StartDate $pendingShiftStart -EndDate $pendingShiftEnd -Segments $segments
        $existingItems += $spItemId
    }

    # End remove items in sharepoint that did not match existing shifts
    $itemsToRemove = $listItems | Where-Object { $_.ID -notin $existingItems }
    foreach($item in $itemsToRemove)
    {
        Write-Debug ("Removing unresolved list item #{0} ({1})" -f $item.ID, $item.Description) 
        $result = Remove-SpListItem @commonParameters -ItemId $item.ID
    }
}
