[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]
    $VmName,
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
$vmToSnapshotList = New-Object -TypeName "System.Collections.ArrayList"
foreach ($queryItem in $vmQueryResult) {
    $vmObject = [PSCustomObject]@{
        name = $queryItem.name
        resourceGroup = $queryItem.resourceGroup
        subscriptionId = $queryItem.subscriptionId
    }
    
    $vmToSnapshotList.Add($vmObject) | Out-Null
}
$currentDate = Get-Date -Format 'yyyy-MM-dd'
$currentTime = Get-Date -Format 'HH-mm-ss'
$vmToSnapshotDiskList = New-Object -TypeName "System.Collections.ArrayList"
#endregion Variables

if (($null -eq $vmToSnapshotList) -or ($vmToSnapshotList.Count -eq 0)) {
    Write-Output "ERROR - No virtual machine yield by Azure Resource Graph"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}

foreach ($vmNameItem in $vmToSnapshotList) {

    if (!((Get-AzContext).Subscription.Id -like $vmNameItem.subscriptionId)) {
        Set-AzContext -SubscriptionId $vmNameItem.subscriptionId
    }
    
    $vmToSnapshot = Get-AzVM -ResourceGroupName $vmNameItem.resourceGroup -Name $vmNameItem.name -ErrorAction SilentlyContinue

    if (($null -eq $vmToSnapshot) -or ($vmToSnapshot.Count -eq 0)) {
        Write-Output "ERROR - Virtual machine $($vmNameItem.name) not found"
    } else {
        Write-Output "INFORMATION - Process to create snapshot for virtual machine $($vmNameItem.name) has started (ticket number $TicketNumber)"
        $vmToSnapshotDiskList.Add($vmToSnapshot.StorageProfile.OsDisk.ManagedDisk.Id) | Out-Null

        if ($vmToSnapshot.StorageProfile.DataDisks.Count -gt 0) {
            $vmToSnapshot.StorageProfile.DataDisks.ManagedDisk.Id | ForEach-Object {$vmToSnapshotDiskList.Add($_) | Out-Null}
        }

        foreach ($snapshotDisk in $vmToSnapshotDiskList) {
            $diskName = $snapshotDisk.Split('/')[-1] + "-snapshot-${currentDate}T${currentTime}"
    
            if ($diskName.Length -gt 80) {
                Write-Output "ERROR - Cannot create snapshot, name ${diskName} is too long"
            }
            else {
                $vmSnapshotConfig = New-AzSnapshotConfig -SkuName Standard_LRS -Location $vmToSnapshot.Location -Tag $vmToSnapshot.Tags -CreateOption Copy -SourceUri $snapshotDisk
                $snapsthoResult = New-AzSnapshot -ResourceGroupName $vmNameItem.resourceGroup -SnapshotName $diskName -Snapshot $vmSnapshotConfig
                Update-AzTag -ResourceId $snapsthoResult.Id -Tag @{"SnapshotTicket" = $TicketNumber} -Operation Merge | Out-Null
                
                if ($snapsthoResult.ProvisioningState -like "Succeeded") {
                    Write-Output "INFORMATION - Snapshot $($snapsthoResult.Name) has been created"
                }
                else {
                    Write-Output "ERROR - Create snapshot for $($vmNameItem.name) failed"
                }
            }        
        }
    }

    $vmToSnapshotDiskList.Clear()
}
Write-Output "INFORMATION - Runbook has finished"
# end of script