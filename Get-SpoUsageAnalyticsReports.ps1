[CmdletBinding()]
Param(
    [String]$siteName,
    [String]$sitePageId
)

function Initialize-MicrosoftIdentityClient {
    ## https://github.com/jpazureid/get-last-signin-reports/blob/master/GetModuleByNuget.ps1 

    $toolDir = "$HOME\spoanalyticstool"
    $sourceNugetExe = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
    $targetNugetExe = "\nuget.exe"

    Write-Log "Start initialization."
    If ( (Test-Path "$toolDir\Microsoft.Identity.Client\Microsoft.Identity.Client.dll" ) -eq $false ){
        Write-Log "Didn't find $toolDir. Start installing Microsoft.Identity.Client."
    
        mkdir $toolDir
        Invoke-WebRequest $sourceNugetExe -OutFile "$toolDir$targetNugetExe" 
        Set-Alias nuget $targetNugetExe -Scope Local -Verbose

        ## Download Microsoft.Identity.Client
        invoke-expression "$toolDir$targetNugetExe install Microsoft.Identity.Client -O $toolDir"
        mkdir "$toolDir\Microsoft.Identity.Client"
        $prtFolder = Get-ChildItem $toolDir | Where-Object {$_.Name -match 'Microsoft.Identity.Client.'}
        Move-Item "$toolDir\$($prtFolder.Name)\lib\net45\*.*" "$toolDir\Microsoft.Identity.Client"
        Remove-Item $toolDir\$($prtFolder.Name) -Force -Recurse
    } else {
        Write-Log "Found $toolDir, so skipped installing Microsoft.Identity.Client."
    }

    Add-Type -Path "$toolDir\Microsoft.Identity.Client\Microsoft.Identity.Client.dll" 

}

Function Write-Log
{
    param(
    [string]$Message,
    [string]$Color = 'White'
    )

    $Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$Date] $Message" -ForegroundColor $Color
}

Function Get-AccessToken() {

    $toolDir = "$HOME\spanalyticstool"
    [string[]]$scopes = @("Sites.Read.All");

    if ($null -eq $local:publicApp) {
        $local:publicApp = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create("14d82eec-204b-4c2f-b7e8-296a70dab67e").WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient").Build();
    }

    $authResult = $local:publicApp.AcquireTokenInteractive($scopes).ExecuteAsync().Result;
    return $authResult
}

Function Get-GraphApi {
    Param(
        [String]$path,
        [String]$token
    )
    
    $baseUrl = "https://graph.microsoft.com/beta"
    $apiUrl = $baseUrl + $path
    Write-Log $apiUrl
    $res = Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $apiUrl -Method GET
    return $res
}

$ErrorActionPreference = "stop"

Initialize-MicrosoftIdentityClient
$token = Get-AccessToken

Write-Log "Caliculate cumulative view" -Color Green

$allContents = New-Object System.Collections.ArrayList
$allContentsWithViewers = New-Object System.Collections.ArrayList
$allContentsWithMonthlyViewers = New-Object System.Collections.ArrayList

# Search your site
$path = "/sites?search=$siteName"
$sites = Get-GraphApi -path $path -token $token.AccessToken
$site = $sites.value | Select-Object webUrl, displayName, id | Out-GridView -PassThru -Title "Please select your site"
$siteId = $site.id

# Get the drive related to your site
$path = "/sites/$siteId/drives/"
$drive = Get-GraphApi -path $path -token $token.AccessToken
$documentDrive = $drive.value | Where-Object { $_.webUrl -like "*Shared%20Documents"}

# Get all files in the document folder.
$skiptoken = $null
Do {

    if ($skiptoken -eq $null){
        $path = "/drives/$($documentDrive.Id)/root/search(q='doc')"
    } else {
        $path = "/drives/$($documentDrive.Id)/root/search(q='doc')" + '?$skiptoken=' + "$skiptoken"
    }
    $files = Get-GraphApi -path $path -token $token.AccessToken
    $files.value | ForEach-Object {
        if ( ("folder" -in $_.PSobject.Properties.Name ) -ne $true ){
            if ($tmp["Name"] -notlike "*.png" -or $tmp["Name"] -notlike "*.jpg"){
                $allContents.add(($_ | Select-Object Name,Id,webUrl)) | Out-Null
            }
        }
    }
    
    $files.'@odata.nextLink' -match "skiptoken=(.*)$" | out-null
    $skiptoken = $Matches[1]        
 
} while ($files.'@odata.nextLink' -ne $null)

# Get all pages in the site page
$skiptoken = $null
Do {

    if ($skiptoken -eq $null){
        $path = "/drives/$sitePageId/search(q='aspx')?$select=name,id,parentReference"
    } else {
        $path = "/drives/$sitePageId/search(q='aspx')?$select=name,id,parentReference" + '?$skiptoken=' + "$skiptoken"
    }
    $pages = Get-GraphApi -path $path -token $token.AccessToken
    $pages.value | ForEach-Object {
        if ( ("folder" -in $_.PSobject.Properties.Name ) -ne $true ){
            $allContents.add(($_ | Select-Object Name,Id,webUrl)) | Out-Null
        }
    }

    $pages.'@odata.nextLink' -match "skiptoken=(.*)$"  | out-null
    $skiptoken = $Matches[1]        
 
} while ($pages.'@odata.nextLink' -ne $null)

$allContents | ForEach {
    $tmp = @{}
    Write-log "Check the analytics of $($_.Name)" -Color Green
    $itemId = $_.id

    if ( $_.webUrl -like "*/Shared%20Documents/*" ){
        $path = "/drives/$($documentDrive.Id)/items/$itemId/analytics/allTime?%24expand=activities(%24filter%3Daccess%20ne%20null)"
    }

    if ( $_.webUrl -like "*/SitePages/*" ){
        $path = "/drives/$sitePageId/items/$itemId/analytics/allTime?%24expand=activities(%24filter%3Daccess%20ne%20null)"
    }

    $activities = Get-GraphApi -path $path -token $token.AccessToken

    $tmp["Name"] = $_.Name
    $tmp["Id"] = $_.Id
    $tmp["webUrl"] = $_.webUrl
    $tmp["actionCount"] = $activities.access.actionCount
    $tmp["actorCount"] = $activities.access.actorCount
    $tmp["Sample 25 Viewers"] = ($activities.activities.actor.user.displayName | Sort-Object ) -join ","

    $allContentsWithViewers.add($tmp) | Out-Null
}

$filename = "spoAnalytics-cumulativeView-$(Get-Date -Format yyyyMMdd-hhmmss).csv"
$allContentsWithViewers = $allContentsWithViewers | ConvertTo-Json -Depth 100 | ConvertFrom-Json
Write-Log "Generate $HOME\spoanalyticstool\$filename" -Color Green
$allContentsWithViewers | Select-Object Name,Id,actionCount,actorCount,"Sample 25 Viewers" | ConvertTo-Csv | Select-Object -skip 1 | Out-File "$HOME\spoanalyticstool\$filename"


$startDateTime = Get-Date (Get-Date -Day 1).AddMonths(-1) -Format yyyy-MM-dd
$endDateTime =  Get-Date (Get-Date -Day 1).AddDays(-1) -Format yyyy-MM-dd

Write-Log "Caliculate the view of the last month from $startDateTime to $endDateTime" -Color Green

$allContents | ForEach {
    $tmp = @{}
    Write-log "Check the getActivitiesByInterval of $($_.Name)" -Color Green
    $itemId = $_.id

    if ( $_.webUrl -like "*/Shared%20Documents/*" ){
        $path = "/drives/$($documentDrive.Id)/items/$itemId/getActivitiesByInterval(startDateTime='$startDateTime',endDateTime='$endDateTime',interval='month')"
    }

    if ( $_.webUrl -like "*/SitePages/*" ){
        $path = "/drives/$sitePageId/items/$itemId/getActivitiesByInterval(startDateTime='$startDateTime',endDateTime='$endDateTime',interval='month')"
    }

    $getActivitiesByInterval = Get-GraphApi -path $path -token $token.AccessToken

    if ($getActivitiesByInterval.value -ne $null){
        $tmp["Name"] = $_.Name
        $tmp["Id"] = $_.Id
        $tmp["webUrl"] = $_.webUrl
        $tmp["startDateTime"] = $getActivitiesByInterval.value.startDateTime
        $tmp["actionCount"] = $getActivitiesByInterval.value.access.actionCount
        $tmp["actorCount"] = $getActivitiesByInterval.value.access.actorCount

        $allContentsWithMonthlyViewers.add($tmp) | Out-Null

    }
}

$filename = "spoAnalytics-MonthlyView-$(Get-Date -Format yyyyMMdd-hhmmss).csv"
$allContentsWithMonthlyViewers = $allContentsWithMonthlyViewers | ConvertTo-Json -Depth 100 | ConvertFrom-Json
Write-Log "Generate $HOME\spoanalyticstool\$filename" -Color Green
$allContentsWithMonthlyViewers | Select-Object startDateTime,Name,Id,actionCount,actorCount | ConvertTo-Csv | Select-Object -skip 1 | Out-File "$HOME\spoanalyticstool\$filename"
