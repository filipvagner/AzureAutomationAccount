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
#   PW0 is 1st Sunday in a month
#   # If PW0 comes out on Saturday or Sunday 1. or 2. then whole schedule in month is moved one week ahead
function Get-PatchWaveZeroDate ($YearToPlan, $DayToPlan, $TimeToStart) {
    
    $patchWaveZeroDatesArray = @()

    for ($i = 1; $i -lt 13; $i++) {
        
        $sundayCounter = 0
    
        for ($j = 1; $j -lt 31; $j++) {        
    
            if((Get-Date -Year $YearToPlan -Month $i -Day $j).DayOfWeek -like 'Sunday') {

                $sundayCounter++

                if (($sundayCounter -eq 1) -and
                 (((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 1) -and ((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 2) -and ((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 3))
                 ) {
                    $pwZschedule = Get-Date -Year $YearToPlan -Month $i -Day $j -Hour $TimeToStart.Hours -Minute $TimeToStart.Minutes -Second $TimeToStart.Seconds
                    $patchWaveZeroDatesArray = $patchWaveZeroDatesArray + $pwZschedule
                    break
                }
                else {
                    $sundayCounter--
                }
            }
        }
    }
    return $patchWaveZeroDatesArray
}

#   PW1 is 2nd Sunday in a month
function Get-PatchWaveOneDate ($YearToPlan, $DayToPlan, $TimeToStart) {
    
    $patchWaveOneDatesArray = @()

    for ($i = 1; $i -lt 13; $i++) {
        
        $sundayCounter = 0
    
        for ($j = 1; $j -lt 31; $j++) {        
    
            if((Get-Date -Year $YearToPlan -Month $i -Day $j).DayOfWeek -like 'Sunday') {

                $sundayCounter++

                if (($sundayCounter -eq 1) -and
                 (((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 1) -and ((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 2) -and ((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 3))
                 ) {
                    $pwOschedule = (Get-Date -Year $YearToPlan -Month $i -Day $j -Hour $TimeToStart.Hours -Minute $TimeToStart.Minutes -Second $TimeToStart.Seconds).AddDays(7)
                    $patchWaveOneDatesArray = $patchWaveOneDatesArray + $pwOschedule
                    break
                }
                else {
                    $sundayCounter--
                }
            }
        }
    }
    return $patchWaveOneDatesArray
}

#   PW2 is 3rd Sunday in a month
function Get-PatchWaveTwoDate ($YearToPlan, $DayToPlan, $TimeToStart) {
    
    $patchWaveTwoDatesArray = @()

    for ($i = 1; $i -lt 13; $i++) {
        
        $sundayCounter = 0
    
        for ($j = 1; $j -lt 31; $j++) {        
    
            if((Get-Date -Year $YearToPlan -Month $i -Day $j).DayOfWeek -like 'Sunday') {

                $sundayCounter++

                if (($sundayCounter -eq 1) -and
                 (((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 1) -and ((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 2) -and ((Get-Date -Year $YearToPlan -Month $i -Day $j).Day -ne 3))
                 ) {
                    $pwDschedule = (Get-Date -Year $YearToPlan -Month $i -Day $j -Hour $TimeToStart.Hours -Minute $TimeToStart.Minutes -Second $TimeToStart.Seconds).AddDays(14)
                    $patchWaveTwoDatesArray = $patchWaveTwoDatesArray + $pwDschedule
                    break
                }
                else {
                    $sundayCounter--
                }
            }
        }
    }
    return $patchWaveTwoDatesArray
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
$dbMandatoryTables = @("linux_patchwave", "linux_timezone", "subscription_permission_check") # Here should be tables that are needed to get information for schedule
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
    SELECT pwname, pwday, pwstart, pwduration, pwisutc, pwregion FROM dbo.linux_patchwave WHERE pwregion = '$pwRegionInDb' AND pwisused = 1;
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
[string[]] $updatesClassification = "Critical", "Security", "Other"
$timeZoneList = @{}
$dbGetQuery = "
SELECT tzcc, tzid FROM dbo.linux_timezone WHERE tzregion = '$pwRegionInDb';
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
    # After change freeze, patching starts always with PW0.
    switch -wildcard ($patchGroupObject.PwName) {
        '*pw0A' {
            Write-Output "INFORMATION - Getting patch wave zero dates"  
            $scheduleStartTimeArray = Get-PatchWaveZeroDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW0 is 1st Sunday in a month"
            break
        }
        '*pw0M' {  
            $scheduleStartTimeArray = Get-PatchWaveZeroDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW0 is 1st Sunday in a month"
            break
        }
        '*pw1A' {
            Write-Output "INFORMATION - Getting patch wave one dates"  
            $scheduleStartTimeArray = Get-PatchWaveOneDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW1 is 2nd Sunday in a month"
            break
        }
        '*pw1M' {
            Write-Output "INFORMATION - Getting patch wave one dates"
            $scheduleStartTimeArray = Get-PatchWaveOneDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW1 is 2nd Sunday in a month"
            break
        }
        '*pw2A'{
            Write-Output "INFORMATION - Getting patch wave two dates"  
            $scheduleStartTimeArray = Get-PatchWaveTwoDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW2 is 3rd Sunday in a month"
            break
        }
        '*pw2M' {
            Write-Output "INFORMATION - Getting patch wave two dates"
            $scheduleStartTimeArray = Get-PatchWaveTwoDate -YearToPlan $yearToPlan -DayToPlan $patchGroupObject.PwDay -TimeToStart $patchGroupObject.PwStart
            $planMonths = 0 .. 10
            $scheduleStartTimeArray = $scheduleStartTimeArray[$planMonths]
            $scheduleDescription = "PW2 is 3rd Sunday in a month"
            break
        }
    }
            
    [Microsoft.Azure.Commands.Automation.Model.Schedule[]] $updateScheduleArray =
    foreach ($scheduleStartTime in $scheduleStartTimeArray) {
        $scheduleDateToString = Get-Date -Date $scheduleStartTime -Format 'yyyy-MM-dd'
        $scheduleName = "Linux-$($patchGroupObject.PwName.Remove(0,4).Replace('_','-'))-$scheduleDateToString"
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
                Linux = $true
                AzureQuery = @($updateMgmtQuery)
                Duration = New-TimeSpan -Hours $patchGroupObject.PwDuration
                RebootSetting =
                if ($patchGroupObject.PwName.EndsWith('M')) {
                    "Never"
                }
                else {
                    "Always"
                }
                IncludedPackageClassification = $updatesClassification
                ResourceGroupName = $scheduleRG
                AutomationAccountName = $scheduleAA
                PreTaskRunbookName = "Backup-AzLinuxRepoFile"
                PreTaskRunbookParameter = @{'PatchWaveTag' = $patchGroupObject.PwName}
                PostTaskRunbookName = "Restore-AzLinuxRepoFile"
                PostTaskRunbookParameter = @{'PatchWaveTag' = $patchGroupObject.PwName}
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
    $patchObjectName = "Linux-" + $patchObjectName
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
            UPDATE dbo.linux_patchwave SET pwisplanned = 1 WHERE pwname = '$($pathcObject.PwName)';
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