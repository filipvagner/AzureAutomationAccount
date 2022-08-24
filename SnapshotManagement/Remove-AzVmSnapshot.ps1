[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]
    $VmName,
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]
    $SnapShotTimeStamp,
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]
    $TicketNumber
)

#region Load modules
$null = Get-Module    -Name 'AzureRM.*' | Remove-Module -Force
$null = Import-Module -Name 'Az.Accounts', 'Az.Automation', 'Az.ResourceGraph'
#endregion Load modules

#region Account connection
$Conn = Get-AutomationConnection -Name AzureRunAsConnection
$ConnectAzAccountParams = @{
  ServicePrincipal      = $true
  Tenant                = $Conn.TenantID
  ApplicationId         = $Conn.ApplicationID
  CertificateThumbprint = $Conn.CertificateThumbprint
}
$null = Connect-AzAccount @ConnectAzAccountParams
$null = Set-AzContext -Subscription (Get-AutomationVariable -Name 'AAccSubscription')
#endregion Account connection

#region Create Azure Resource Graph Query and yield machines
$argQuery = [System.Text.StringBuilder]::new()
$argQuery.Append("resources ") | Out-Null
$argQuery.Append("| where type =~ 'Microsoft.Compute/virtualMachines' ") | Out-Null
$argQuery.Append("| where ") | Out-Null

foreach ($vmItem in $VmName.Split(',')) {
    $vmItemToAppend = $vmItem.Trim()
    $argQuery.Append("name has ""$vmItemToAppend"" or ") | Out-Null
}
$argQuery.Remove(($argQuery.Length - 4), 4) | Out-Null
$argQuery.Append(" | sort by subscriptionId") | Out-Null
$argQuery.Append(" | project name, resourceGroup, subscriptionId") | Out-Null

$vmQueryResult = Search-AzGraph -Query $argQuery.ToString()
if (($null -eq $vmQueryResult) -or ($vmQueryResult.Count -eq 0)) {
    Write-Output "ERROR - No virtual machine yield by Azure Resource Graph"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}
#endregion Create Azure Resource Graph Query and yield machines

#region Variables
$vmToRemoveSnapshotList = New-Object -TypeName "System.Collections.ArrayList"
foreach ($queryItem in $vmQueryResult) {
    $vmObject = [PSCustomObject]@{
        name = $queryItem.name
        resourceGroup = $queryItem.resourceGroup
        subscriptionId = $queryItem.subscriptionId
    }
    
    $vmToRemoveSnapshotList.Add($vmObject) | Out-Null
}
#endregion Variables

if (($null -eq $vmToRemoveSnapshotList) -or ($vmToRemoveSnapshotList.Count -eq 0)) {
    Write-Output "ERROR - No virtual machine yield by Azure Resource Graph"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}

foreach ($vmNameItem in $vmToRemoveSnapshotList) {

    if (!((Get-AzContext).Subscription.Id -like $vmNameItem.subscriptionId)) {
        Set-AzContext -SubscriptionId $vmNameItem.subscriptionId
    }

    Write-Output "INFORMATION - Process to remove snapshot for virtual machine $($vmNameItem.name) has started (ticket number $TicketNumber)"
    $snapshotToRemoveList = Get-AzSnapshot -ResourceGroupName $vmNameItem.resourceGroup | Where-Object {($_.Name.StartsWith($vmNameItem.name)) -and ($_.Name.EndsWith($SnapShotTimeStamp))}
    
    if (($snapshotToRemoveList.Count -eq 0) -or ($null -eq $snapshotToRemoveList)) {
        Write-Output "WARNING - No snapshot for virtual machine $($vmNameItem.name) found"
    }
    else {
        foreach ($snapshotToRemove in $snapshotToRemoveList) {
            $removeResult = Remove-AzSnapshot -ResourceGroupName $vmNameItem.resourceGroup -SnapshotName $snapshotToRemove.Name -Force
        
            if ($removeResult.Status -like "Succeeded") {
                Write-Output "INFORMATION - Snapshot $($snapshotToRemove.Name) has been removed"
            }
            else {
                Write-Output "ERROR - Remove snapshot $($snapshotToRemove.Name) failed"
            }

            $removeResult = $null
        }
    }
}
Write-Output "INFORMATION - Runbook has finished"
# end of script