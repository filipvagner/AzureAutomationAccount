# Script's parameters
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [ValidateSet('Linux', 'Windows', 'ActiveDirectory', 'All')]
    [string]
    $OperatingSystem,
    [Parameter()]
    [AllowNull()] # ISO country code, All
    [string]
    $CountryCode,
    [Parameter()]
    [AllowNull()] # pw code, All
    [string]
    $PatchWaveCode,
    [bool]
    $DeleteAll = $false
)

# Module operations
$null = Get-Module    -Name 'AzureRM.*' | Remove-Module -Force
$null = Import-Module -Name 'Az.Accounts', 'Az.Automation', 'SqlServer'

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
# Module operations end

$scheduleRg = Get-AutomationVariable -Name 'AAccGroup'
$scheduleAa = Get-AutomationVariable -Name 'AAccName'
$pwRegionInDb = Get-AutomationVariable -Name 'pwRegion'
$listOfDeploymentSchedules = Get-AzAutomationSoftwareUpdateConfiguration -ResourceGroupName $scheduleRG -AutomationAccountName $scheduleAA
$filteredDeploymentSchedules = New-Object -TypeName "System.Collections.ArrayList"
$dsSucceeded = $true
$dsScopeToRemove = ""
$dbConParams = @{
    ServerInstance = Get-AutomationVariable -Name 'scheduleServerInstance'
    Database = Get-AutomationVariable -Name 'scheduleDatabase'
    Username = Get-AutomationVariable -Name 'scheduleDbUsername'
    Password = Get-AutomationVariable -Name 'scheduleDbPwd'
    Query = ""
}

# Check database connection and availability
$dbMandatoryTables = @("linux_patchwave", "windows_patchwave", "activedirectory_patchwave")
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

# Scope to remove
if (($OperatingSystem -like "All") -and ($CountryCode -like "All") -and ($PatchWaveCode -like "All") -and ($DeleteAll)) {
    $dsScopeToRemove = "RemoveAllDeploymentSchedules"
}
elseif (($OperatingSystem -notlike 'All') -and !([string]::IsNullOrEmpty($OperatingSystem)) -and ([string]::IsNullOrEmpty($CountryCode)) -and ([string]::IsNullOrEmpty($PatchWaveCode)) -and !($DeleteAll)) {
    $dsScopeToRemove = "RemoveByOs"
}
elseif (!([string]::IsNullOrEmpty($OperatingSystem)) -and (![string]::IsNullOrEmpty($CountryCode)) -and ([string]::IsNullOrEmpty($PatchWaveCode)) -and !($DeleteAll)) {
    $dsScopeToRemove = "RemoveByOsAndCountry"
}
elseif (!([string]::IsNullOrEmpty($OperatingSystem)) -and ([string]::IsNullOrEmpty($CountryCode)) -and (![string]::IsNullOrEmpty($PatchWaveCode)) -and !($DeleteAll)) {
    $dsScopeToRemove = "RemoveByOsAndPatchWave"
}
elseif (!([string]::IsNullOrEmpty($OperatingSystem)) -and (![string]::IsNullOrEmpty($CountryCode)) -and (![string]::IsNullOrEmpty($PatchWaveCode)) -and !($DeleteAll)) {
    $dsScopeToRemove = "RemoveByOsAndCountryAndPatchWave"
}
else {
    $dsScopeToRemove = "InvalidParameters"
    $dsSucceeded = $false
}
Write-Output "INFORMATION - Scope to remove has been set to $dsScopeToRemove"
# Scope to remove end

# Check if any deployemnt schedules were found
if (!([string]::IsNullOrEmpty($OperatingSystem)) -and ($OperatingSystem -notlike "All") -and ($dsScopeToRemove -notlike "InvalidParameters")) {
    foreach ($deploymentSchedule in $listOfDeploymentSchedules) {
        if ($deploymentSchedule.Name.StartsWith($OperatingSystem)) {
            $filteredDeploymentSchedules.Add($deploymentSchedule)
        }
    }
    $deploymentSchedule = $null
    if (($listOfDeploymentSchedules.Count -eq 0) -or ($filteredDeploymentSchedules.Count -eq 0)) {
        $dsSucceeded = $false
    }
}
# Check if any deployemnt schedules were found end

# Remove deployment schedules in Azure
switch ($dsScopeToRemove) {
    "RemoveByOs" {
        if (!$dsSucceeded) {
            Write-Output "WARNING - No deployment schedule found"
            break
        }
        foreach ($deploymentSchedule in $filteredDeploymentSchedules) {
            Write-Output "INFORMATION - Removing deployment schedule $($deploymentSchedule.Name)"
            Remove-AzAutomationSoftwareUpdateConfiguration -Name $deploymentSchedule.Name -ResourceGroupName $scheduleRG -AutomationAccountName $scheduleAA
        }
        break
    }
    "RemoveByOsAndCountry" {
        if (!$dsSucceeded -or (($filteredDeploymentSchedules | Where-Object {$_.Name -like "*$CountryCode*"}).Count -eq 0)) {
            Write-Output "WARNING - No deployment schedule found"
            $dsSucceeded = $false
            break
        }
        foreach ($deploymentSchedule in $filteredDeploymentSchedules | Where-Object {$_.Name -like "*$CountryCode*"}) {
            Write-Output "INFORMATION - Removing deployment schedule $($deploymentSchedule.Name)"
            Remove-AzAutomationSoftwareUpdateConfiguration -Name $deploymentSchedule.Name -ResourceGroupName $scheduleRG -AutomationAccountName $scheduleAA
        }
        break
    }
    "RemoveByOsAndPatchWave" {
        if (!$dsSucceeded -or (($filteredDeploymentSchedules | Where-Object {$_.Name -like "*$PatchWaveCode*"}).Count -eq 0)) {
            Write-Output "WARNING - No deployment schedule found"
            $dsSucceeded = $false
            break
        }
        foreach ($deploymentSchedule in $filteredDeploymentSchedules | Where-Object {$_.Name -like "*$PatchWaveCode*"}) {
            Write-Output "INFORMATION - Removing deployment schedule $($deploymentSchedule.Name)"
            Remove-AzAutomationSoftwareUpdateConfiguration -Name $deploymentSchedule.Name -ResourceGroupName $scheduleRG -AutomationAccountName $scheduleAA
        }
        break
    }
    "RemoveByOsAndCountryAndPatchWave" {
        if (!$dsSucceeded -or (($filteredDeploymentSchedules | Where-Object {($_.Name -like "*$CountryCode*") -and ($_.Name -like "*$PatchWaveCode*")}).Count -eq 0)) {
            Write-Output "WARNING - No deployment schedule found"
            $dsSucceeded = $false
            break
        }
        foreach ($deploymentSchedule in $filteredDeploymentSchedules | Where-Object {($_.Name -like "*$CountryCode*") -and ($_.Name -like "*$PatchWaveCode*")}) {
            Write-Output "INFORMATION - Removing deployment schedule $($deploymentSchedule.Name)"
            Remove-AzAutomationSoftwareUpdateConfiguration -Name $deploymentSchedule.Name -ResourceGroupName $scheduleRG -AutomationAccountName $scheduleAA
        }
        break
    }
    "RemoveAllDeploymentSchedules" {
        if (!$dsSucceeded -or ($listOfDeploymentSchedules.Count -eq 0)) {
            Write-Output "WARNING - No deployment schedule found"
            $dsSucceeded = $false
            break
        }
        foreach ($deploymentSchedule in $listOfDeploymentSchedules) {
            Write-Output "INFORMATION - Removing deployment schedule $($deploymentSchedule.Name)"
            Remove-AzAutomationSoftwareUpdateConfiguration -Name $deploymentSchedule.Name -ResourceGroupName $scheduleRG -AutomationAccountName $scheduleAA
        }
        break
    }
    "InvalidParameters" {
        Write-Output "ERROR - Some parameter did not meet condition to remove deployment schedules"
        $dsSucceeded = $false
        break
    }
    Default {
        Write-Ouput "ERROR - Some parameter did not meet condition to remove deployment schedules"
        $dsSucceeded = $false
        break
    }
}
# Remove deployment schedules in Azure end

if (!$dsSucceeded) {
    Write-Output "INFORMATION - No changes made in database"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}

# Update records in database
$osInQuery = $OperatingSystem.ToLower()
switch ($dsScopeToRemove) {
    "RemoveByOs" {
        $dbUpdateQuery = "UPDATE dbo.$($osInQuery)_patchwave SET pwisplanned = 0 WHERE pwregion = '$pwRegionInDb';"
        $dbConParams.Query = $dbUpdateQuery
        Invoke-Sqlcmd @dbConParams
        $dbUpdateQuery = "UPDATE dbo.$($osInQuery)_patchwave SET pwisused = 0 WHERE pwregion = '$pwRegionInDb';"
        $dbConParams.Query = $dbUpdateQuery
        Invoke-Sqlcmd @dbConParams
        Write-Output "INFORMATION - All records in database for $OperatingSystem in region $pwRegionInDb has been updated"
        break
    }
    "RemoveByOsAndCountry" {
        $dbUpdateQuery = [System.Text.StringBuilder]::new()
        $null = $dbUpdateQuery.Append("UPDATE dbo.$($osInQuery)_patchwave SET pwisplanned = 0 WHERE pwname LIKE '%")
        $null = $dbUpdateQuery.Append($CountryCode)
        $null = $dbUpdateQuery.Append("%' AND pwregion = '$pwRegionInDb';")
        $dbConParams.Query = $dbUpdateQuery.ToString()
        Invoke-Sqlcmd @dbConParams
        $null = $dbUpdateQuery.Clear()
        $null = $dbUpdateQuery.Append("UPDATE dbo.$($osInQuery)_patchwave SET pwisused = 0 WHERE pwname LIKE '%")
        $null = $dbUpdateQuery.Append($CountryCode)
        $null = $dbUpdateQuery.Append("%' AND pwregion = '$pwRegionInDb';")
        $dbConParams.Query = $dbUpdateQuery.ToString()
        Invoke-Sqlcmd @dbConParams
        Write-Output "INFORMATION - All records in database for $OperatingSystem with country code $CountryCode has been updated"
        break
    }
    "RemoveByOsAndPatchWave" {
        $dbUpdateQuery = [System.Text.StringBuilder]::new()
        $null = $dbUpdateQuery.Append("UPDATE dbo.$($osInQuery)_patchwave SET pwisplanned = 0 WHERE pwname LIKE '%")
        $null = $dbUpdateQuery.Append($PatchWaveCode)
        $null = $dbUpdateQuery.Append("%' AND pwregion = '$pwRegionInDb';")
        $dbConParams.Query = $dbUpdateQuery.ToString()
        Invoke-Sqlcmd @dbConParams
        $null = $dbUpdateQuery.Clear()
        $null = $dbUpdateQuery.Append("UPDATE dbo.$($osInQuery)_patchwave SET pwisused = 0 WHERE pwname LIKE '%")
        $null = $dbUpdateQuery.Append($PatchWaveCode)
        $null = $dbUpdateQuery.Append("%' AND pwregion = '$pwRegionInDb';")
        $dbConParams.Query = $dbUpdateQuery.ToString()
        Invoke-Sqlcmd @dbConParams
        Write-Output "INFORMATION - All records in database for $OperatingSystem with patch wave $PatchWaveCode has been updated"
        break
    }
    "RemoveByOsAndCountryAndPatchWave" {
        $dbUpdateQuery = [System.Text.StringBuilder]::new()
        $null = $dbUpdateQuery.Append("UPDATE dbo.$($osInQuery)_patchwave SET pwisplanned = 0 WHERE pwname LIKE '%")
        $null = $dbUpdateQuery.Append($CountryCode)
        $null = $dbUpdateQuery.Append("%' AND pwname LIKE '%")
        $null = $dbUpdateQuery.Append($PatchWaveCode)
        $null = $dbUpdateQuery.Append("%' AND pwregion = '$pwRegionInDb';")
        $dbConParams.Query = $dbUpdateQuery.ToString()
        Invoke-Sqlcmd @dbConParams
        $null = $dbUpdateQuery.Clear()
        $dbUpdateQuery = [System.Text.StringBuilder]::new()
        $null = $dbUpdateQuery.Append("UPDATE dbo.$($osInQuery)_patchwave SET pwisused = 0 WHERE pwname LIKE '%")
        $null = $dbUpdateQuery.Append($CountryCode)
        $null = $dbUpdateQuery.Append("%' AND pwname LIKE '%")
        $null = $dbUpdateQuery.Append($PatchWaveCode)
        $null = $dbUpdateQuery.Append("%' AND pwregion = '$pwRegionInDb';")
        $dbConParams.Query = $dbUpdateQuery.ToString()
        Invoke-Sqlcmd @dbConParams
        Write-Output "INFORMATION - All records in database for $OperatingSystem with country code $CountryCode and with patch wave $PatchWaveCode has been updated"
        break
    }
    "RemoveAllDeploymentSchedules" {
        # Update Active Direcotry table
        $dbUpdateQuery = "UPDATE dbo.activedirectory_patchwave SET pwisplanned = 0 WHERE pwregion = '$pwRegionInDb';"
        $dbConParams.Query = $dbUpdateQuery
        Invoke-Sqlcmd @dbConParams
        $dbUpdateQuery = "UPDATE dbo.activedirectory_patchwave SET pwisused = 0 WHERE pwregion = '$pwRegionInDb';"
        $dbConParams.Query = $dbUpdateQuery
        Invoke-Sqlcmd @dbConParams
        # Update Linux table
        $dbUpdateQuery = "UPDATE dbo.linux_patchwave SET pwisplanned = 0 WHERE pwregion = '$pwRegionInDb';"
        $dbConParams.Query = $dbUpdateQuery
        Invoke-Sqlcmd @dbConParams
        $dbUpdateQuery = "UPDATE dbo.linux_patchwave SET pwisused = 0 WHERE pwregion = '$pwRegionInDb';"
        $dbConParams.Query = $dbUpdateQuery
        Invoke-Sqlcmd @dbConParams
        # Update Windows table
        $dbUpdateQuery = "UPDATE dbo.windows_patchwave SET pwisplanned = 0 WHERE pwregion = '$pwRegionInDb';"
        $dbConParams.Query = $dbUpdateQuery
        Invoke-Sqlcmd @dbConParams
        $dbUpdateQuery = "UPDATE dbo.windows_patchwave SET pwisused = 0 WHERE pwregion = '$pwRegionInDb';"
        $dbConParams.Query = $dbUpdateQuery
        Invoke-Sqlcmd @dbConParams
        Write-Output "INFORMATION - All records in database in region $pwRegionInDb has been updated"
        break
    }
    "InvalidParameters" {
        Write-Output "ERROR - Some parameter did not meet condition to remove deployment schedules"
        $dsSucceeded = $false
        break
    }
    Default {
        Write-Ouput "ERROR - Some parameter did not meet condition to remove deployment schedules"
        $dsSucceeded = $false
        break
    }
}
# Update records in database end

if (!$dsSucceeded) {
    Write-Output "INFORMATION - No changes made in database"
    Write-Output "INFORMATION - Runbook has stopped"
}
else {
    Write-Output "INFORMATION - Runbook has finished"
}
# end of script