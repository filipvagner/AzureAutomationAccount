[CmdletBinding()]
param (
    [bool]
    $ReadSubscription = $false,
    [bool]
    $PermissionDifferenceCheck = $false
)

#region Module operations
$null = Get-Module    -Name 'AzureRM.*' | Remove-Module -Force
$null = Import-Module -Name 'Az.Accounts', 'Az.Automation', 'SqlServer'
#endregion Module operations

#region Automation Account Data and connection
$Conn = Get-AutomationConnection -Name AzureRunAsConnection
$ConnectAzAccountParams = @{
  ServicePrincipal      = $true
  Tenant                = $Conn.TenantID
  ApplicationId         = $Conn.ApplicationID
  CertificateThumbprint = $Conn.CertificateThumbprint
}
$null = Connect-AzAccount @ConnectAzAccountParams
$null = Set-AzContext -Subscription (Get-AutomationVariable -Name 'AAccSubscription')
$jobId = $PSPrivateMetadata.Values.Guid
#endregion Automation Account Data and connection

#region Check database connection and availability
$dbConParams = @{
    ServerInstance = Get-AutomationVariable -Name 'scheduleServerInstance'
    Database = Get-AutomationVariable -Name 'scheduleDatabase'
    Username = Get-AutomationVariable -Name 'scheduleDbUsername'
    Password = Get-AutomationVariable -Name 'scheduleDbPwd'
    Query = ""
}
$dbMandatoryTables = @("subscription_permission_check") # Here should be tables that are needed to get information for schedule
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
#endregion Check database connection and availability

#region variables
$arReader = "<Application (client) ID>"
[string[]] $arDeployment = "<Application (client) ID>", "<Application (client) ID>", "<Application (client) ID>" # APAC, EMEA, NAM
[string[]] $arOperation = "<Application (client) ID>", "<Application (client) ID>", "<Application (client) ID>" # APAC, EMEA, NAM
$aaId = (Get-AzContext).Account.Id
$aaRg = Get-AutomationVariable -Name 'AAccGroup'
$aaName = Get-AutomationVariable -Name 'AAccName'
#endregion variables

#region reader account
if (($aaId -like $arReader) -and $ReadSubscription -and !($PermissionDifferenceCheck)) {
    Write-Output "INFORMATION - Update for read account $aaId in progress"

    $subscriptionAzureList = Get-AzSubscription | Select-Object -Property Name, Id, State
    $dbSelectQuery = "
    SELECT subscriptionid FROM dbo.subscription_permission_check;
    "
    $dbConParams.Query = $dbSelectQuery
    $subscriptionDbList = Invoke-Sqlcmd @dbConParams

    foreach ($subscriptionItem in $subscriptionAzureList) {
        if (($null -eq $subscriptionDbList) -or (!$subscriptionDbList.subscriptionid.Contains($subscriptionItem.Id))) {
            if ($subscriptionItem.State -like "Enabled") {
                Write-Output "INFORMATION - Subscription $($subscriptionItem.Name) - $($subscriptionItem.Id) updated in database"

                $subRegion = $subscriptionItem.Name.Split('.')[0]                
                $dbInsertQuery = "
                INSERT INTO dbo.subscription_permission_check (subscriptionid, subscriptionname, subscriptionregion) VALUES ('$($subscriptionItem.Id)', '$($subscriptionItem.Name)', '$subRegion');
                "
                $dbConParams.Query = $dbInsertQuery
                Invoke-Sqlcmd @dbConParams
            }
        }
    }
    $dbConParams.Query = ""
    $subscriptionDbList = $null

    Write-Output "INFORMATION - Update in database completed"
}
elseif (($aaId -like $arReader) -and !($ReadSubscription) -and $PermissionDifferenceCheck) {
    Write-Output "INFORMATION - Checking for permissions differences in database"

    $dbSelectQuery = "
    SELECT subscriptionname FROM dbo.subscription_permission_check WHERE aadeploymenthasaccess = 1 AND (aaoperationhasaccess = 0 OR aaoperationhasaccess IS NULL) AND excludesubscription != 1;
    "
    $dbConParams.Query = $dbSelectQuery
    $subscriptionDbList = Invoke-Sqlcmd @dbConParams

    if(!($null -eq $subscriptionDbList)) {
        foreach ($subscriptionItem in $subscriptionDbList) {
            Write-Output "WARNING - Permission difference found on subscription $($subscriptionItem.subscriptionname)"
        }

        #region Send email
        Write-Output "INFORMATION - Sending results by email"
        $runbookResultEmailParams = @{
            ToEmailAddress = "name@domain"
            EmailSubject = "Check Subscriptions Permission Report"
            LogAnalyticsWorskspace = $true
            RunbookName = "Update-AzSubscriptionPermissionDatabase"
            JobId = $jobId.ToString()
            AllMessages = $true
        }

        Start-AzAutomationRunbook `
            -AutomationAccountName $aaName `
            -Name 'Send-RunbookResultByEmail' `
            -ResourceGroupName $aaRg `
            -Parameters $runbookResultEmailParams
        #endregion Send email
    }
    else {
        Write-Output "INFORMATION - No permission difference found in database"
    }
}
elseif (($aaId -like $arReader) -and $ReadSubscription -and $PermissionDifferenceCheck) {
    Write-Output "ERROR - Read account can run read subscription mode only or permission difference check mode only"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}

if (!($aaId -like $arReader) -and ($ReadSubscription -or $PermissionDifferenceCheck)) {
    Write-Output "ERROR - This automation account cannot run in read subscription or permissions check mode"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}
#endregion reader account

#region deployemnt account
if (($arDeployment.Contains($aaId)) -and !($ReadSubscription -or $PermissionDifferenceCheck)) {
    Write-Output "INFORMATION - Update for deployment account $aaId in progress"

    $aaRegion = $aaName.Split('-')[1]
    $dbSelectQuery = "
    SELECT subscriptionid FROM dbo.subscription_permission_check WHERE subscriptionregion = '$aaRegion';
    "
    $dbConParams.Query = $dbSelectQuery
    $subscriptionDbList = Invoke-Sqlcmd @dbConParams

    if ($null -eq $subscriptionDbList) {
        Write-Output "ERROR - No records found in database"
        Write-Output "INFORMATION - Runbook has stopped"
        exit
    }

    foreach ($subscriptionItem in $subscriptionDbList) {
        $subscriptionId = $subscriptionItem.subscriptionid
        $roleStatus = Get-AzRoleAssignment -ServicePrincipalName $aaId -Scope "/subscriptions/$subscriptionId"
        $aaRole = $roleStatus.RoleDefinitionName
        
        if (!($null -eq $roleStatus)) {
            Write-Output "INFORMATION - Subscription $subscriptionId updated in database"

            $dbUpdateQuery = "
            UPDATE dbo.subscription_permission_check SET aadeploymentname = '$aaName', aadeploymenthasaccess = 1, aadeploymentrole = '$aaRole' WHERE subscriptionid = '$subscriptionId'; 
            "
            $dbConParams.Query = $dbUpdateQuery
            Invoke-Sqlcmd @dbConParams
        }
        $roleStatus = $null
    }
    $dbConParams.Query = ""
    $subscriptionDbList = $null

    Write-Output "INFORMATION - Update in database completed"
}
#endregion deployemnt account

#region operation account
if (($arOperation.Contains($aaId)) -and !($ReadSubscription -or $PermissionDifferenceCheck)) {
    Write-Output "INFORMATION - Update for operation account $aaId in progress"

    $aaRegion = $aaName.Split('-')[1]
    $dbSelectQuery = "
    SELECT subscriptionid FROM dbo.subscription_permission_check WHERE aadeploymenthasaccess = 1 AND subscriptionregion = '$aaRegion';
    "
    $dbConParams.Query = $dbSelectQuery
    $subscriptionDbList = Invoke-Sqlcmd @dbConParams

    if ($null -eq $subscriptionDbList) {
        Write-Output "ERROR - No records found in database"
        Write-Output "INFORMATION - Runbook has stopped"
        exit
    }

    foreach ($subscriptionItem in $subscriptionDbList) {
        $subscriptionId = $subscriptionItem.subscriptionid
        $roleStatus = Get-AzRoleAssignment -ServicePrincipalName $aaId -Scope "/subscriptions/$subscriptionId"
        $aaRole = $roleStatus.RoleDefinitionName
        
        if (!($null -eq $roleStatus)) {
            Write-Output "INFORMATION - Subscription $subscriptionId updated in database"

            $dbUpdateQuery = "
            UPDATE dbo.subscription_permission_check SET aaoperationname = '$aaName', aaoperationhasaccess = 1, aaoperationrole = '$aaRole' WHERE subscriptionid = '$subscriptionId'; 
            "
            $dbConParams.Query = $dbUpdateQuery
            Invoke-Sqlcmd @dbConParams
        }
        $roleStatus = $null
    }
    $dbConParams.Query = ""
    $subscriptionDbList = $null

    Write-Output "INFORMATION - Update in database completed"
}
#endregion operation account