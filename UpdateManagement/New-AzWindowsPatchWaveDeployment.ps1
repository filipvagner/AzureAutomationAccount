# Module operations
$null = Get-Module    -Name 'AzureRM.*' | Remove-Module -Force
$null = Import-Module -Name 'Az.Accounts', 'Az.Automation', 'SqlServer'
# Module operations end

# Automation Account Data and connection
$Conn = Get-AutomationConnection -Name AzureRunAsConnection
$ConnectAzAccountParams = @{
  ServicePrincipal      = $true
  Tenant                = $Conn.TenantID
  ApplicationId         = $Conn.ApplicationID
  CertificateThumbprint = $Conn.CertificateThumbprint
}
$null = Connect-AzAccount @ConnectAzAccountParams
$null = Set-AzContext -Subscription (Get-AutomationVariable -Name 'AAccSubscription')
# Automation Account Data and connection end

# Functions
function Get-PatchWaveDate ($YearToPlan, $DayToPlan, $TimeToStart, $PatchWave) {
    
    $patchWaveDatesArray = @()
    switch ($PatchWave) {
        "pwT" { $tuesdayCounter = 1; break }
        "pw0" { $tuesdayCounter = 2; break }
        "pw1" { $tuesdayCounter = 3; break }
        "pw2" { $tuesdayCounter = 4; break }
        Default { throw "Unknown patch wave specified (use 'pwT', 'pw0', 'pw1' or 'pw2')" }
    }

    for ($monthCounter = 1; $monthCounter -lt 13; $monthCounter++) {
        $dayCounter = 0
        $dayToAdd = 0

        do {
            $dayCounter++
            
            if((Get-Date -Year $YearToPlan -Month $monthCounter -Day $dayCounter).DayOfWeek -like "Tuesday") {
                $tuesdayToStart = (Get-Date -Year $YearToPlan -Month $monthCounter -Day $dayCounter).AddDays(7 * $tuesdayCounter)
                break
            }
            
        } while ($true)

        do {
            $dayToAdd++
            if ($tuesdayToStart.AddDays($dayToAdd).DayOfWeek.ToString().StartsWith($DayToPlan, $true, [System.Globalization.CultureInfo]::InvariantCulture)) {
                $pwDate = (Get-Date -Year $tuesdayToStart.Year -Month $tuesdayToStart.Month -Day $tuesdayToStart.Day -Hour $TimeToStart.Hours -Minute $TimeToStart.Minutes -Second $TimeToStart.Seconds).AddDays($dayToAdd)
                $patchWaveDatesArray = $patchWaveDatesArray + $pwDate
                break
            }
        } while ($true)
    }

    return $patchWaveDatesArray
}

function Get-PatchWaveTimeZone ($PatchWaveName, $timeZoneList) {
    $regexCheck = [System.Text.RegularExpressions.Regex]::new("_..._")

    return $timeZoneList[$regexCheck.Match($PatchWaveName).Value.Replace('_','').Trim()]
}
# Functions end

# Check database connection and availability
$dbConParams = @{
    ServerInstance = Get-AutomationVariable -Name 'scheduleServerInstance'
    Database = Get-AutomationVariable -Name 'scheduleDatabase'
    Username = Get-AutomationVariable -Name 'scheduleDbUsername'
    Password = Get-AutomationVariable -Name 'scheduleDbPwd'
    Query = ""
}
$dbMandatoryTables = @("windows_patchwave", "windows_timezone", "subscription_permission_check") # Here should be tables that are needed to get information for schedule
$dbTableCounter = 0
if ((Test-NetConnection -ComputerName $dbConParams.ServerInstance -Port 1433).TcpTestSucceeded) {
    Write-Output "INFORMATION - Connection to database successful"
} 
else {
    Write-Output "ERROR - Connection to database failed"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}

$dbGetQuery = "SELECT * FROM INFORMATION_SCHEMA.TABLES;"
$dbConParams.Query = $dbGetQuery
$dbTableList = Invoke-Sqlcmd @dbConParams
foreach ($table in $dbTableList) {
    if ($dbMandatoryTables.Contains($table.Table_Name)) {
        Write-Output "INFORMATION - Table $($table.Table_Name) found"
        $dbTableCounter++
    }   
}
if ($dbTableCounter -eq $dbMandatoryTables.Count) {
    Write-Output "INFORMATION - All mandatory tables in database found"
} else {
    Write-Output "ERROR - Some mandatory table is missing"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}
$dbConParams.Query = ""
# Check database connection and availability end

# Variables
## Common variables
$pwRegionInDb = Get-AutomationVariable -Name 'pwRegion'
$dbGetQuery = "
    SELECT pwname, pwday, pwstart, pwduration, pwisutc, pwregion FROM dbo.windows_patchwave WHERE pwregion = '$pwRegionInDb' AND pwisused = 1;
    "
$dbConParams.Query = $dbGetQuery
$patchGroupList = Invoke-Sqlcmd @dbConParams
$dbConParams.Query = ""
if (($null -eq $patchGroupList) -or ($patchGroupList.Count -eq 0)) {
    Write-Output "WARNING - No patch wave to be scheduled found"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}

$patchGroupObjectList = New-Object -TypeName "System.Collections.ArrayList"
$patchGroupList | Foreach-Object {
    $pwObj = [PSCustomOBject]@{
        PwName = $_.pwname
        PwDay = $_.pwday
        PwStart = $_.pwstart
        PwDuration = $_.pwduration
        PwIsUtc =
        if ($_.pwisutc) {
            $true
        }
        else {
            $false
        }
        PwRegion = $_.pwregion
    }
    $patchGroupObjectList.Add($pwObj)
}
[int]$yearToPlan = Get-AutomationVariable -Name 'scheduleYear'
$currentDateUtc = (Get-Date).ToUniversalTime()
$scheduleRg = Get-AutomationVariable -Name 'AAccGroup'
$scheduleAa = Get-AutomationVariable -Name 'AAccName'
[string[]] $updatesClassification = "Critical", "Security"
$timeZoneList = @{}
$dbGetQuery = "
SELECT tzcc, tzid FROM dbo.windows_timezone WHERE tzregion = '$pwRegionInDb';
    "
$dbConParams.Query = $dbGetQuery
Invoke-Sqlcmd @dbConParams | ForEach-Object {$timeZoneList.Add($_.tzcc, $_.tzid)}
$dbConParams.Query = ""

$dbGetQuery = "
SELECT CONCAT('/subscriptions/', subscriptionid) AS subscriptionId FROM dbo.subscription_permission_check WHERE subscriptionregion = '$pwRegionInDb' AND ((aadeploymenthasaccess = 1 AND aaoperationhasaccess = 1) AND excludesubscription != 1);
    "
$dbConParams.Query = $dbGetQuery
$queryScope = Invoke-Sqlcmd @dbConParams
$dbConParams.Query = ""
# Variables end

# Deployment schedule creation
Write-Output "INFORMATION - Deployment schedule has started"
foreach ($patchGroupObject in $patchGroupObjectList) {
    Write-Output "INFORMATION - Creating deployment schedules for $($patchGroupObject.PwName)"
    #FIXME switch statement does not work when more conditions is used in one case!!!

    # During change freeze patching is not performed.
    # Change freeze starts in December (from each array last patch wave is removed).
    # After change freeze, patching starts always with PWT.
    switch -wildcard ($patchGroupObject.PwName) {
        '*pwT' {
            Write-Output "INFORMATION - Getting patch wave test dates"  
            $scheduleStartTimeArray = Get-PatchWaveDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart -PatchWave "pwT"
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PWT is 1st day to plan (pwday) after 2nd TUE"
            break
        }
        '*pw0A' {
            Write-Output "INFORMATION - Getting patch wave zero dates"  
            $scheduleStartTimeArray = Get-PatchWaveDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart -PatchWave "pw0"
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW0 is 1st day to plan (pwday) after 3rd TUE"
            break
        }
        '*pw1A' {
            Write-Output "INFORMATION - Getting patch wave one dates"  
            $scheduleStartTimeArray = Get-PatchWaveDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart -PatchWave "pw1"
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW1 is 1st day to plan (pwday) after 4th TUE"
            break
        }
        '*pw1M' {
            Write-Output "INFORMATION - Getting patch wave one dates"
            $scheduleStartTimeArray = Get-PatchWaveDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart -PatchWave "pw1"
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW1 is 1st day to plan (pwday) after 4th TUE"
            break
        }
        '*pw2A'{
            Write-Output "INFORMATION - Getting patch wave two dates"  
            $scheduleStartTimeArray = Get-PatchWaveDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart -PatchWave "pw2"
            $planMonths = 0 .. 9
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW2 is 1st day to plan (pwday) after 5th TUE"
            break
        }
        '*pw2M' {
            Write-Output "INFORMATION - Getting patch wave two dates"
            $scheduleStartTimeArray = Get-PatchWaveDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart -PatchWave "pw2"
            $planMonths = 0 .. 9
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW2 is 1st day to plan (pwday) after 5th TUE"
            break
        }
    }
            
    [Microsoft.Azure.Commands.Automation.Model.Schedule[]] $updateScheduleArray =
    foreach ($scheduleStartTime in $scheduleStartTimeArray) {
        $scheduleDateToString = Get-Date -Date $scheduleStartTime -Format 'yyyy-MM-dd'
        $scheduleName = "Windows-$($patchGroupObject.PwName.Remove(0,4).Replace('_','-'))-$scheduleDateToString"
        $updateScheduleParams = @{
            Name = $scheduleName
            StartTime = $scheduleStartTime
            Description  = $scheduleDescription
            OneTime = $true
            TimeZone =
            if ($patchGroupObject.PwIsUtc) {
                "Etc/UTC"
            } else {
                Get-PatchWaveTimeZone -PatchWaveName $patchGroupObject.PwName -TimeZoneList $timeZoneList
            }
            ForUpdateConfiguration = $true
            ResourceGroupName = $scheduleRG
            AutomationAccountName = $scheduleAA
        }
        New-AzAutomationSchedule @updateScheduleParams
    }

    $updateMgmtQueryParams = @{
        Scope = @($queryScope | Select-Object -ExpandProperty subscriptionId)
        Location = @(
            switch ($patchGroupObject.PwRegion) {
                'apac' { "japaneast"; break }
                'emea' { "westeurope"; break }
                'nam' { "eastus"; break }
            }
        )
        Tag = @{
            PatchWave = $patchGroupObject.PwName
        }
        # FilterOperator = "Any"
        ResourceGroupName = $scheduleRG
        AutomationAccountName = $scheduleAA
                
    }
    $updateMgmtQuery = New-AzAutomationUpdateManagementAzureQuery @updateMgmtQueryParams

    foreach ($updateSchedule in $updateScheduleArray) {
        if ($updateSchedule.StartTime.UtcDateTime -gt $currentDateUtc) {
            $pwtUpdateConfigParams = @{
                Schedule = $updateSchedule
                Windows = $true
                AzureQuery = @($updateMgmtQuery)
                Duration = New-TimeSpan -Hours $patchGroupObject.PwDuration
                RebootSetting =
                if ($patchGroupObject.PwName.EndsWith('M')) {
                    "Never"
                }
                else {
                    "Always"
                }
                IncludedUpdateClassification = $updatesClassification
                ResourceGroupName = $scheduleRG
                AutomationAccountName = $scheduleAA
            }
            New-AzAutomationSoftwareUpdateConfiguration @pwtUpdateConfigParams
            Write-Output "INFORMATION - Deployment schedule $($updateSchedule.Name) created"
        }
    }
}
# Deployment schedule creation end

# Check if all update schedules were successful
# If yes, then patch wave is updated in database as 'is planned'
Write-Output "INFORMATION - Deployement schedules check has started"
foreach ($pathcObject in $patchGroupObjectList) {
    Write-Output "INFORMATION - Checking deployment schedules for patch wave $($pathcObject.PwName)"
    $scheduleUpdateSuccessful = $true
    $patchObjectName = $pathcObject.PwName.Remove(0,4).Replace('_', '-')
    $patchObjectName = "Windows-" + $patchObjectName
    $updateConfigurationList = Get-AzAutomationSoftwareUpdateConfiguration -ResourceGroupName $scheduleRg -AutomationAccountName $scheduleAa | Where-Object {$_.Name -match $patchObjectName}
    foreach ($updateConfigurationItem in $updateConfigurationList) {
        if ($updateConfigurationItem.ProvisioningState -notlike "Succeeded") {
            $scheduleUpdateSuccessful = $false
            Write-Output "WARNING - Deployment schedule $($updateConfigurationItem.Name) did not succeeded"
        }
    }
    if ($scheduleUpdateSuccessful) {
        $scheduleList = Get-AzAutomationSchedule -ResourceGroupName $scheduleRg -AutomationAccountName $scheduleAa | Where-Object {$_.Name.StartsWith($patchObjectName)}
        foreach ($scheduleItem in $scheduleList) {
            if ($scheduleItem.IsEnabled) {
                Write-Output "INFORMATION - Schedule $($scheduleItem.Name) is already set as enabled"
            }
            else {
                Write-Output "INFORMATION - Enabling schedule $($scheduleItem.Name)"
                Set-AzAutomationSchedule -Name $scheduleItem.Name -ResourceGroupName $scheduleRg -AutomationAccountName $scheduleAa -IsEnabled $true
            }
        }
        # It is better after trying to enable each schedule to do check again because error can occur while sending the request
        $scheduleList.Clear()
        $scheduleList = Get-AzAutomationSchedule -ResourceGroupName $scheduleRg -AutomationAccountName $scheduleAa | Where-Object {$_.Name.StartsWith($patchObjectName)}
        foreach ($scheduleItem in $scheduleList) {
            if (!($scheduleItem.IsEnabled)) {
                Write-Output "WARNING - Schedule $($scheduleItem.Name) is not enabled"
            }
        }
        Write-Output "INFORMATION - All deployment schedules for $($pathcObject.PwName) succeeded"
        $dbUpdateQuery = "
            UPDATE dbo.windows_patchwave SET pwisplanned = 1 WHERE pwname = '$($pathcObject.PwName)';
            "
            $dbConParams.Query = $dbUpdateQuery
            Invoke-Sqlcmd @dbConParams
            $dbConParams.Query = ""
            Write-Output "INFORMATION - Patch wave $($pathcObject.PwName) record in database updated"
    } else {
        Write-Output "ERROR - Check deployment schedules for patch wave $($pathcObject.PwName)"
        Write-Output "INFORMATION - Runbook has stopped"
        exit
    }
}
# Check if all update schedules were successful end