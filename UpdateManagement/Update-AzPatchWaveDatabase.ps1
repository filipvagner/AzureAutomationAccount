# Module operations
$null = Get-Module    -Name 'AzureRM.*' | Remove-Module -Force
$null = Import-Module -Name 'Az.Accounts', 'Az.Automation', 'Az.ResourceGraph', 'SqlServer'
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

# Check database connection and availability
$dbConParams = @{
    ServerInstance = Get-AutomationVariable -Name 'scheduleServerInstance'
    Database = Get-AutomationVariable -Name 'scheduleDatabase'
    Username = Get-AutomationVariable -Name 'scheduleDbUsername'
    Password = Get-AutomationVariable -Name 'scheduleDbPwd'
    Query = $dbGetQuery
}
$dbMandatoryTables = @("linux_patchwave", "windows_patchwave", "activedirectory_patchwave") # Here should be tables that are needed to get information for schedule
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
$argQuery = "
resources
| where type =~ 'Microsoft.Compute/virtualMachines'
| mvexpand pwos = properties.storageProfile.osDisk.osType
| mvexpand pwname = tags.PatchWave
| where isnotnull(pwname)
| project pwname, pwos
"

$pwQueryResult = Search-AzGraph -Query $argQuery -First 1000
if (($null -eq $pwQueryResult) -or ($pwQueryResult.Count -eq 0)) {
    Write-Output "WARNING - No patch wave tag returned by Azure Resource Graph"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}
$pwAdObjectList = New-Object -TypeName "System.Collections.ArrayList"
$pwLinuxObjectList = New-Object -TypeName "System.Collections.ArrayList"
$pwWindowsObjectList = New-Object -TypeName "System.Collections.ArrayList"
$regexCheckPwAd = [System.Text.RegularExpressions.Regex]::new("GRP_ADS_pw\d[AM]_g\d", "ignoreCase")
# Variables end

# Update patch wave records in database
Write-Output "INFORMATION - Update patch wave records in database has started"
foreach ($pwItem in $pwQueryResult) {
    $pwObject = [PSCustomOBject]@{
        PwName = $pwItem.pwname
        PwOs = $pwItem.pwos
    }
    if ($regexCheckPwAd.IsMatch($pwObject.PwName)) {
        if (($pwAdObjectList.Count -eq 0) -or (!$pwAdObjectList.PwName.Contains($pwObject.PwName))) {
            $null = $pwAdObjectList.Add($pwObject)
            $dbGetQuery = "
            DECLARE @PwName varchar(255) = '$($pwObject.PwName)'
            SELECT (CASE WHEN EXISTS(SELECT 1 FROM dbo.activedirectory_patchwave WITH(NOLOCK) WHERE pwname = @PwName)
            THEN '1'
            ELSE '0' END) AS [recordexist]
            "
            $dbConParams.Query = $dbGetQuery
            if ((Invoke-Sqlcmd @dbConParams).recordexist -like 1) {
                $dbConParams.Query = ""
                $dbUpdateQuery = "
                UPDATE dbo.activedirectory_patchwave SET pwisused = 1 WHERE pwname = '$($pwObject.PwName)';
                "
                $dbConParams.Query = $dbUpdateQuery
                Invoke-Sqlcmd @dbConParams
                $dbConParams.Query = ""
                Write-Output "INFORMATION - Patch wave $($pwObject.PwName) record in database updated"
            }
            else {
                Write-Output "WARNING - Patch wave $($pwObject.PwName) record does not exist in database"
            }
        }
    }
    else {
        switch ($pwObject.PwOs) {
            'Linux' { 
                if (($pwLinuxObjectList.Count -eq 0) -or (!$pwLinuxObjectList.PwName.Contains($pwObject.PwName))) {
                    $null = $pwLinuxObjectList.Add($pwObject)
                    $dbGetQuery = "
                    DECLARE @PwName varchar(255) = '$($pwObject.PwName)'
                    SELECT (CASE WHEN EXISTS(SELECT 1 FROM dbo.linux_patchwave WITH(NOLOCK) WHERE pwname = @PwName)
                    THEN '1'
                    ELSE '0' END) AS [recordexist]
                    "
                    $dbConParams.Query = $dbGetQuery
                    if ((Invoke-Sqlcmd @dbConParams).recordexist -like 1) {
                        $dbConParams.Query = ""
                        $dbUpdateQuery = "
                        UPDATE dbo.linux_patchwave SET pwisused = 1 WHERE pwname = '$($pwObject.PwName)';
                        "
                        $dbConParams.Query = $dbUpdateQuery
                        Invoke-Sqlcmd @dbConParams
                        $dbConParams.Query = ""
                        Write-Output "INFORMATION - Patch wave $($pwObject.PwName) record in database updated"
                    }
                    else {
                        Write-Output "WARNING - Patch wave $($pwObject.PwName) record does not exist in database"
                    }
                }
                break
            }
            'Windows' { 
                if (($pwWindowsObjectList.Count -eq 0) -or (!$pwWindowsObjectList.PwName.Contains($pwObject.PwName))) {
                    $null = $pwWindowsObjectList.Add($pwObject)
                    $dbGetQuery = "
                    DECLARE @PwName varchar(255) = '$($pwObject.PwName)'
                    SELECT (CASE WHEN EXISTS(SELECT 1 FROM dbo.windows_patchwave WITH(NOLOCK) WHERE pwname = @PwName)
                    THEN '1'
                    ELSE '0' END) AS [recordexist]
                    "
                    $dbConParams.Query = $dbGetQuery
                    if ((Invoke-Sqlcmd @dbConParams).recordexist -like 1) {
                        $dbConParams.Query = ""
                        $dbUpdateQuery = "
                        UPDATE dbo.windows_patchwave SET pwisused = 1 WHERE pwname = '$($pwObject.PwName)';
                        "
                        $dbConParams.Query = $dbUpdateQuery
                        Invoke-Sqlcmd @dbConParams
                        $dbConParams.Query = ""
                        Write-Output "INFORMATION - Patch wave $($pwObject.PwName) record in database updated"
                    }
                    else {
                        Write-Output "WARNING - Patch wave $($pwObject.PwName) record does not exist in database"
                    }
                }
                break
            }
        }
    }
}
# Update patch wave records in database end