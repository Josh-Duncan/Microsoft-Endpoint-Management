# Last tested with MECM Current Branch 2002
# It is unclear how much backwards compatibility it may have.  Use at your own peril
#
# This script needs to be run with an accoun that has at least the "Read-Only Analyst"
# This script does not currently need write permissions to anything.
#
# Josh Duncan
# 2020-04-24
# https://github.com/Josh-Duncan
# https://github.com/Josh-Duncan/MEM_MECM/wiki/Quick-Health-Check.ps1
#
# TODO: Add optiona logging output with rolling option
# TODO: Update this to trigger multiple updates at the same time and monitor.  Will need to closely monitor server performance.


$TeamsChannelWebhook = ""
    # Define the Teams Channel Webhook address

$SUMinCompliance = 70
    # Set the compliance percentage you want to configure as "out of compliant"

$SUComplianceSched = 14
    # Number of days from deadline you want to hide non compliant deployments

$MaxSummarizationMinutes = 120
    # the summarization age that is acceptible for the report

$ScriptDebug = $false
    # Doesn't do much, but it does output the JSON message to the console.

$MessageTitle = "MECM System Status"
    # Pretty Self explanitory

$SiteCode = "Primary Site Code"
    # Site code is only needed when multiple sites are involved, or if the powershell module hasn't been loaded

$ProviderMachineName = "Primary Server Name"
    # Primary Site Server Name

# ------------------------------------------------
# No variables to edit below here...

$DateNow = Get-Date

if ($ScriptDebug -eq $true)
{
    Write-Host ""
    Write-Host "Scritp is currently running in Debug Mode" -ForegroundColor Magenta
    Write-Host ""

    Write-Host "Script start time:" -ForegroundColor Magenta
    Write-Host " |       Script Run Time:"$DateNow -ForegroundColor Magenta
    Write-Host " | (UTC) Script Run Time:"$DateNow.ToUniversalTime() -ForegroundColor Magenta
    Write-Host ""
}

# ------------------------------------------------
# Validate site information

try
{
    $CMSite = Get-CMSite
}
catch
{
    Write-Host ""
    Write-Host "Configuring Modules..."

    try
    {
        if ($ScriptDebug -eq $true){Write-Host " | Setting Execution Policy..." -ForegroundColor Magenta}
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process | Out-Null
        
        if ($ScriptDebug -eq $true){Write-Host " | Importing CM Module" -ForegroundColor Magenta}
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"   | Out-Null
        
        if ($ScriptDebug -eq $true){Write-Host " | Setting PS Drive..." -ForegroundColor Magenta}
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName  | Out-Null
        
        if ($ScriptDebug -eq $true){Write-Host " | Setting Location..." -ForegroundColor Magenta}
        Set-Location "$($SiteCode):\" | Out-Null
    }
    Catch
    {
        Write-Host ""
        Write-Error "Failed to Import PS Module"
        Write-host "Script exited early" -ForegroundColor Red
        Exit
    }

    try
    {
        Write-Host ""
        if ($ScriptDebug -eq $true){Write-Host ("Connecting to Site Code " + $SiteCode + "...") -ForegroundColor Magenta}
        $CMSite = Get-CMSite -SiteCode $SiteCode 
        if ($ScriptDebug -eq $true){Write-Host (Write-Host " | "$CMsite.Count" Sites Found") -ForegroundColor Magenta}
    }
    catch
    {
        Write-Host ""
        Write-host "Error: Failed to connect to any site, please check site variables or run from CM enabled ISE" -ForegroundColor Red
        Write-host "Script exited early" -ForegroundColor Red
        exit
    }
}

if ($CMSite.Count -ne 1)
{
    Write-Host ""
    Write-host ("Error: Unsupported number of sites found (" + $CMSite.Count + ")") -ForegroundColor Red
    Write-host "Script exited early" -ForegroundColor Red
    exit
}

# ------------------------  Do Stuff ------------------------

Write-Host ""
Write-Host $MessageTitle
Write-host ($CMSite.SiteName + " - " + $CMSite.SiteCode + " (v." + $CMSite.Version + ")")

$JSON_ActivityTitle = $MessageTitle
$JSON_ActivitySubTitle = ($CMSite.SiteName + " - " + $CMSite.SiteCode + " (v." + $CMSite.Version + ")")

if ($TeamsChannelWebhook -eq "")
{
    Write-Host ""
    Write-Warning "Webhook not configured"
    write-host ""
}

# ------------------------------------------------
# Display Active Alert Information

$CMAlerts_Critical = get-CMAlert | where {(($_.Severity -eq 1) -and ($_.AlertState -eq 0))}
$CMAlerts_Warning = get-CMAlert | where {(($_.Severity -eq 2) -and ($_.RuleState -eq 0))}

write-host ""
write-host "Active Alerts"
Write-Host (" | Critical: " + $CMAlerts_Critical.count)
Write-Host (" | Warning:  " + $CMAlerts_Warning.count)

$JSON_CriticalAlert = "{""name"": ""Critical Alerts"",""value"": """ + $CMAlerts_Critical.count + """},"

Write-Host ""
write-host "   Critical"

if ($CMAlerts_Critical.count -eq 0)
    {Write-Host "    * None"}
else
{
    foreach ($Alert in $CMAlerts_Critical)
    {
        Write-Host "    | "$Alert.Name
        $JSON_CriticalAlert = $JSON_CriticalAlert + "{""name"": """",""value"": "">" + $Alert.Name + """},"
    }
}
$Alert = @()

Write-Host ""
write-host "   Warning"
if ($CMAlerts_Warning.count -eq 0)
    {Write-Host "    * None"}
else
{
    foreach ($Alert in $CMAlerts_Warning)
        {Write-Host "    | "$Alert.Name}
}

# ------------------------------------------------
# Display Software Update Information

$BelowCompliance = 0
$SUDeployments = Get-CMDeployment | where {($_.FeatureType -eq 5)}

Write-Host ""
write-host "Checking Update Deployment Summarization times"

$i = 0

# TODO: Update this to trigger multiple updates at the same time and monitor.  Will need to closely monitor server performance.

foreach ($SUDeployment in $SUDeployments)
{
    Write-Progress -Activity "Updating Deployment Summarization..." -Status "Status:" -PercentComplete (($i/$SUDeployments.Count)*100) -CurrentOperation ($i.ToString() + "/" + $SUDeployments.Count + " processed")

    if ($SUDeployment.SummarizationTime.AddMinutes($MaxSummarizationMinutes) -lt $DateNow.ToUniversalTime())
    {
        Write-Host "  + Updating:"$SUDeployment.ApplicationName -ForegroundColor Yellow
        Invoke-CMDeploymentSummarization -DeploymentId $SUDeployment.DeploymentID
        if ($ScriptDebug -eq $true)
        {
            Write-Host "    |        Now:"$DateNow -ForegroundColor Magenta
            Write-Host "    | (UTC)  Now:"$DateNow.ToUniversalTime() -ForegroundColor Magenta
            Write-Host "    | (UTC) Sync:"$SUDeployment.SummarizationTime -ForegroundColor Magenta
        }
        # The lesson to be leared in all this UTC debugging is that CM returns UTC as the summarization time by default. 
        do
        {
            Start-Sleep  -Seconds 5
            $DeploymentTimeCheck = Get-CMDeployment -DeploymentId $SUDeployment.DeploymentID | Select-Object SummarizationTime
            if($ScriptDebug -eq $true){Write-Host "    || (UTC) Sync:"$SUDeployment.SummarizationTime -ForegroundColor Magenta}
        
        } while ($DeploymentTimeCheck.SummarizationTime.AddMinutes($MaxSummarizationMinutes) -lt $DateNow.ToUniversalTime())
        
        $DeploymentTimeCheck = @()
    }
    else
    {
        Write-Host " OK Skipping:"$SUDeployment.ApplicationName -ForegroundColor Green
        if ($ScriptDebug -eq $true)
        {
            Write-Host "    |        Now:"$DateNow -ForegroundColor Magenta
            Write-Host "    | (UTC)  Now:"$DateNow.ToUniversalTime() -ForegroundColor Magenta
            Write-Host "    | (UTC) Sync:"$SUDeployment.SummarizationTime -ForegroundColor Magenta
        }
    }
    $i++
}    

Write-Progress -Activity "Updating Deployment Summarization..." -Status "Status:" -PercentComplete (($i/$SUDeployments.Count)*100) -CurrentOperation ($i.ToString() + "/" + $SUDeployments.Count + " processed") -Complete

$SUDeployments = Get-CMDeployment | where {($_.FeatureType -eq 5)}

$i = 0

foreach ($SUDeployment in $SUDeployments)
{
    if (([math]::Round($SUDeployment.NumberSuccess / $SUDeployment.NumberTargeted * 100,1)) -lt $SUMinCompliance)
    {$BelowCompliance++}
}

write-host ""
write-host "Software Update Deployments: "
write-host " | Total:    " $SUDeployments.Count
write-host " | Below $SUMinCompliance%:" $BelowCompliance

$JSON_SUCompliance = "{""name"": ""Software Updates"",""value"": ""$BelowCompliance deployments below $SUMinCompliance% compliance""},"

if ($SUDeployments.Count -ne 0)
{
    foreach ($SUDeployment in $SUDeployments)
    {
        if (([math]::Round($SUDeployment.NumberSuccess / $SUDeployment.NumberTargeted * 100,1)) -lt $SUMinCompliance)
        {
            $DeploymentCompliance_Output = ([math]::Round($SUDeployment.NumberSuccess / $SUDeployment.NumberTargeted * 100,1))
            $diff = New-TimeSpan -Start $SUDeployment.EnforcementDeadline.AddDays(+$SUComplianceSched).ToString("yyyy-MM-dd") -End $DateNow
            
            write-host ("    | " + $DeploymentCompliance_Output + "% " + $SUDeployment.ApplicationName + " (" + $diff.Days + " days past deadline)")
            if ($SUDeployment.EnforcementDeadline.AddDays(+$SUComplianceSched).ToString("yyyy-MM-dd") -lt $DateNow.ToString("yyyy-MM-dd"))
            {$JSON_SUCompliance = $JSON_SUCompliance + "{""name"": """",""value"": "">" + $DeploymentCompliance_Output + "% " `
                + $SUDeployment.ApplicationName + " (" + $diff.Days + " days past deadline)""},"}
        }
    }
}

$CMSiteUpdates = get-cmsiteupdate -Fast | where {($_.State -ne 196612)} 

Write-Host ""
Write-Host "CM Updates Available: "$CMSiteUpdates.count

$JSON_CMUpdate = "{""name"": ""CM Updates"",""value"": """ + $CMSiteUpdates.count + " updates available""},"

if ($CMSiteUpdates.count -ne 0)
{
    foreach ($CMUPdate in $CMSiteUpdates)
    {
        Write-Host " | "$CMUPdate.Name" ("$CMUPdate.FullVersion")"
        $JSON_CMUpdate = $JSON_CMUpdate + ("{""name"": """",""value"": "">" + $CMUPdate.Name + " (" +$CMUPdate.FullVersion + ")" + """},")
    }
}

# ------------------------------------------------
# Build and send the JSON formatted message

$JSON_body = @"
{
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    "themeColor": "0076D7",
    "summary": "Status Updates",
    "sections": [{
        "activityTitle": "$JSON_ActivityTitle",
        "activitySubtitle": "$JSON_ActivitySubTitle",
        "facts":
        [
            $JSON_CriticalAlert
            $JSON_CMUpdate
            $JSON_SUCompliance
        ],
        "markdown": true
    }],

}
"@

if ($ScriptDebug -eq $true)
{
    write-host "--- DEBUG - JSON Message Body ---" -ForegroundColor Magenta
    write-host $JSON_body  -ForegroundColor Magenta
}
if ($TeamsChannelWebhook -ne "")
{
    try
    {
        Invoke-RestMethod -Uri $TeamsChannelWebHook -Method Post -Body $JSON_body -ContentType 'application/json' | Out-Null
    }
    catch
    {
        write-host ""
        Write-host "Error Sending Message:" -ForegroundColor Red
        Write-host $_.exception.message -ForegroundColor Red
    }
}
else
{
    Write-host ""
    Write-warning "Webhook not configured.  Message not sent"
    try
    {
        ConvertFrom-Json $JSON_body
    }
    catch
    {
        Write-Warning "Could not convert Json output"
    }
}
if ($ScriptDebug -eq $true)
{
    Write-Host ""
    Write-Host "Script ran to completion" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "Script end time:" -ForegroundColor Magenta
    Write-Host " |       Script Run Time:"$DateNow -ForegroundColor Magenta
    Write-Host " | (UTC) Script Run Time:"$DateNow.ToUniversalTime() -ForegroundColor Magenta
    Write-Host ""
}

