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
    $TicketNumber,
    [bool]
    $RestoreWithSameName = $false
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
$vmToRestoreList = New-Object -TypeName "System.Collections.ArrayList"
foreach ($queryItem in $vmQueryResult) {
    $vmObject = [PSCustomObject]@{
        name = $queryItem.name
        resourceGroup = $queryItem.resourceGroup
        subscriptionId = $queryItem.subscriptionId
    }
    
    $vmToRestoreList.Add($vmObject) | Out-Null
}
$vmToRestoreDiskList = New-Object -TypeName "System.Collections.ArrayList"
#endregion Variables

if (($null -eq $vmToRestoreList) -or ($vmToRestoreList.Count -eq 0)) {
    Write-Output "ERROR - No virtual machine yield by Azure Resource Graph"
    Write-Output "INFORMATION - Runbook has stopped"
    exit
}

foreach ($vmNameItem in $vmToRestoreList) {

    if (!((Get-AzContext).Subscription.Id -like $vmNameItem.subscriptionId)) {
        Set-AzContext -SubscriptionId $vmNameItem.subscriptionId
    }

    Write-Output "INFORMATION - Process to restore virtual machine $($vmNameItem.name) from snapshot has started (ticket number $TicketNumber)"
    $vmToRestore = Get-AzVM -ResourceGroupName $vmNameItem.resourceGroup -Name $vmNameItem.name
    Write-Output "INFORMATION - Stopping virtual machine $($vmNameItem.name)"
    $vmToRestore | Stop-AzVM -Force -NoWait | Out-Null
    $stopVmCounter = 0
    
    do {
        Write-Output "INFORMATION - Stopping virtual machine $($vmNameItem.name) in progress"
        $vmToRestoreStopResult = Get-AzVM -ResourceGroupName $vmToRestore.ResourceGroupName -Name $vmToRestore.Name -Status
        if ($stopVmCounter -eq 20) {
            $vmDeallocatedSuccess = $false
            break
        }
        Start-Sleep -Seconds 15
        $stopVmCounter++
        $vmDeallocatedSuccess = $true
    } until (($vmToRestoreStopResult.Statuses[1].Code -like "PowerState/deallocated") -and ($vmToRestoreStopResult.Statuses[1].DisplayStatus -like "VM deallocated"))

    $restoreVmSuccess = $true
    $restoreDiskCoutner = 1

    if (!$vmDeallocatedSuccess) {
        Write-Output "ERROR - Virtual machine $($vmNameItem.name) failed to stop"
        Write-Output "WARNING - Restore for virtual machine $($vmNameItem.name) skipped"
        $vmDeallocatedSuccess = $null
    }
    else {
        Write-Output "INFORMATION - Virtual machine $($vmNameItem.name) has been stopped and deallocated"

        Write-Output "INFORMATION - Getting virtual machine disks info"
        $vmToRestoreDiskList.Add($vmToRestore.StorageProfile.OsDisk.ManagedDisk.Id) | Out-Null
        
        if ($vmToRestore.StorageProfile.DataDisks.Count -gt 0) {
            $vmToRestore.StorageProfile.DataDisks.ManagedDisk.Id | ForEach-Object {$vmToRestoreDiskList.Add($_) | Out-Null}
        }

        Write-Output "INFORMATION - Getting virtual machine snapshots info"
        $snapshotForRestoreList = Get-AzSnapshot -ResourceGroupName $vmToRestore.ResourceGroupName | Where-Object {($_.Name.StartsWith($vmToRestore.name)) -and ($_.Name.EndsWith($SnapShotTimeStamp))}

        #region Restore disk with same name
        if ($RestoreWithSameName -and ($snapshotForRestoreList.Count -gt 0)) {
            Write-Output "INFORMATION - Restoring disk with same name"
            $rgNameForRestore = "rg-restore_" + $vmNameItem.name + "_${TicketNumber}-${SnapShotTimeStamp}"
            $rgForRestore = New-AzResourceGroup -Name $rgNameForRestore -Location $vmToRestore.Location -Tag $vmToRestore.Tags
            
            if ($rgForRestore.ProvisioningState -like "Succeeded") {
                Write-Output "INFORMATION - Resource group for restore $($rgForRestore.ResourceGroupName) has been created"
                Update-AzTag -ResourceId $rgForRestore.ResourceId -Tag @{"RestoreTicket" = $TicketNumber} -Operation Merge | Out-Null
                
                foreach ($restoreDiskItem in $vmToRestoreDiskList) {
                    
                    if (!$restoreVmSuccess) {
                        break
                    }

                    $snapshotToRestore = $snapshotForRestoreList | Where-Object {$_.CreationData.SourceResourceId -like $restoreDiskItem}
                    $originalDisk = Get-AzDisk -ResourceGroupName $vmToRestore.ResourceGroupName -DiskName $restoreDiskItem.Split('/')[-1]
                    $restoreDiskConfig = New-AzDiskConfig -SkuName $originalDisk.Sku.Name -Location $originalDisk.Location -Tag $originalDisk.Tags -CreateOption Copy -SourceUri $snapshotToRestore.Id -Zone $originalDisk.Zones
                    $restoreDisk = New-AzDisk -Disk $restoreDiskConfig -ResourceGroupName $rgForRestore.ResourceGroupName -DiskName $originalDisk.Name
                    
                    if ($restoreDisk.ProvisioningState -like "Succeeded") {
                        Write-Output "INFORMATION - New disk $($restoreDisk.Name) from snapshot $($snapshotToRestore.Name) has been created"

                        if (($vmToRestoreDiskList[0] -like $restoreDiskItem) -and !($null -eq $originalDisk.OsType)) {
                            Write-Output "INFORMATION - Setting operating system disk"
                            Set-AzVMOSDisk -VM $vmToRestore -ManagedDiskId $restoreDisk.Id -Name $restoreDisk.Name | Out-Null
                            $updateResult = Update-AzVM -ResourceGroupName $vmToRestore.ResourceGroupName -VM $vmToRestore
                        }                        
                        else {
                            Write-Output "INFORMATION - Setting data disk"
                            $originalDiskLun = $vmToRestore.StorageProfile.DataDisks | ForEach-Object {if($_.ManagedDisk.Id -like $snapshotToRestore.CreationData.SourceResourceId) {$_.Lun}}
                            Remove-AzVMDataDisk -VM $vmToRestore -DataDiskNames $originalDisk.Name | Out-Null
                            $updateResult = Update-AzVM -ResourceGroupName $vmToRestore.ResourceGroupName -VM $vmToRestore
                            if ($updateResult.IsSuccessStatusCode) {
                                Write-Output "INFORMATION - Original disk $($originalDisk.Name) has been dettached from virtual machine $($vmNameItem.name)"

                                Add-AzVMDataDisk -VM $vmToRestore -ManagedDiskId $restoreDisk.Id -Name $restoreDisk.Name -CreateOption Attach -Lun $originalDiskLun | Out-Null
                                $updateResult = Update-AzVM -ResourceGroupName $vmToRestore.ResourceGroupName -VM $vmToRestore
                                if (!$updateResult.IsSuccessStatusCode) {
                                    $restoreVmSuccess = $false
                                }
                            }
                            else {
                                Write-Output "ERROR - Dettach original disk $($originalDisk.Name) from virtual machine $($vmNameItem.name) failed"
                                $restoreVmSuccess = $false
                            }                            
                        }

                        if ($updateResult.IsSuccessStatusCode) {
                            Write-Output "INFORMATION - Restored disk $($restoreDisk.Name) has been attached to virtual machine $($vmNameItem.name)"

                            $originalDiskRemoveResult = Remove-AzDisk -ResourceGroupName $vmToRestore.ResourceGroupName -DiskName $originalDisk.Name -Force
                            if ($originalDiskRemoveResult.Status -like "Succeeded") {
                                Write-Output "INFORMATION - Disk with ID $($originalDisk.Id) was deleted"
                                $diskToMove = Get-AzResource -ResourceId $restoreDisk.Id
                                Move-AzResource -DestinationResourceGroupName $vmToRestore.ResourceGroupName -ResourceId $diskToMove.ResourceId -Force
                                $vmToRestore = Get-AzVM -ResourceGroupName $vmNameItem.resourceGroup -Name $vmNameItem.name
                                Write-Output "INFORMATION - Restored disk has been moved to resource group $($vmToRestore.ResourceGroupName)"
                                
                                if ($restoreDiskCoutner -eq $vmToRestoreDiskList.Count) {
                                    $rgForRestoreRemoveResult = Remove-AzResourceGroup -Id $rgForRestore.ResourceId -Force
                                    
                                    if ($rgForRestoreRemoveResult) {
                                        Write-Output "INFORMATION - Resource group $($rgForRestore.ResourceGroupName) has been deleted"
                                    }
                                    else {
                                        Write-Output "WARNING - Resource group $($rgForRestore.ResourceGroupName) was not deleted"
                                        Write-Output "INFORMATION - Check status of $($rgForRestore.ResourceGroupName) and delete it manualy"
                                    }
                                }
                                else {
                                    $restoreDiskCoutner++
                                }
                            }
                            else {
                                Write-Output "ERROR - Delete original disk $($originalDisk.Id) failed"
                                $restoreVmSuccess = $false
                            }
                        }
                        else {
                            $restoreVmSuccess = $false
                        }
                    }
                    else {
                        Write-Output "ERROR - Restore from snapshot $($snapshotToRestore.Name) failed"
                        $restoreVmSuccess = $false
                    }

                    if (!$updateResult.IsSuccessStatusCode) {
                        Write-Output "ERROR - Attach disk $($restoreDisk.Name) to virtual machine $($vmNameItem.name) failed"
                        $restoreVmSuccess = $false
                    }

                    $snapshotToRestore = $null
                    $originalDiskLun = $null
                }
            }
            else {
                Write-Output "ERROR - Create resource group for restore $($rgForRestore.ResourceGroupName) failed"
                $restoreVmSuccess = $false
            }
        }
        #endregion Restore disk with same name

        #region Restore disk with new name
        elseif (!$RestoreWithSameName -and ($snapshotForRestoreList.Count -gt 0)){
            Write-Output "INFORMATION - Restoring disks with new name"
            foreach ($restoreDiskItem in $vmToRestoreDiskList) {

                if (!$restoreVmSuccess) {
                    break
                }

                $snapshotToRestore = $snapshotForRestoreList | Where-Object {$_.CreationData.SourceResourceId -like $restoreDiskItem}
                $originalDisk = Get-AzDisk -ResourceGroupName $vmToRestore.ResourceGroupName -DiskName $restoreDiskItem.Split('/')[-1]
                $restoreDiskConfig = New-AzDiskConfig -SkuName $originalDisk.Sku.Name -Location $originalDisk.Location -Tag $originalDisk.Tags -CreateOption Copy -SourceUri $snapshotToRestore.Id -Zone $originalDisk.Zones
                $restoreDiskName = $originalDisk.Name + "-restore-" + $SnapShotTimeStamp
                
                if (Get-AzDisk -ResourceGroupName $vmToRestore.ResourceGroupName -DiskName $restoreDiskName -ErrorAction SilentlyContinue) {
                    Write-Output "WARNING - Disk with name $restoreDiskName already exist"
                    $restoreVmSuccess = $false
                }
                else {
                    $restoreDisk = New-AzDisk -Disk $restoreDiskConfig -ResourceGroupName $vmToRestore.ResourceGroupName -DiskName $restoreDiskName

                    if ($restoreDisk.ProvisioningState -like "Succeeded") {
                        Write-Output "INFORMATION - New disk $($restoreDisk.Name) from snapshot $($snapshotToRestore.Name) has been created"
                        
                        if (($vmToRestoreDiskList[0] -like $restoreDiskItem) -and !($null -eq $originalDisk.OsType)) {
                            Write-Output "INFORMATION - Setting operating system disk"
                            Set-AzVMOSDisk -VM $vmToRestore -ManagedDiskId $restoreDisk.Id -Name $restoreDisk.Name | Out-Null
                            $updateResult = Update-AzVM -ResourceGroupName $vmToRestore.ResourceGroupName -VM $vmToRestore
                        }                        
                        else {
                            Write-Output "INFORMATION - Setting data disk"
                            $originalDiskLun = $vmToRestore.StorageProfile.DataDisks | ForEach-Object {if($_.ManagedDisk.Id -like $snapshotToRestore.CreationData.SourceResourceId) {$_.Lun}}
                            Remove-AzVMDataDisk -VM $vmToRestore -DataDiskNames $originalDisk.Name | Out-Null
                            $updateResult = Update-AzVM -ResourceGroupName $vmToRestore.ResourceGroupName -VM $vmToRestore
                            
                            if ($updateResult.IsSuccessStatusCode) {
                                Write-Output "INFORMATION - Original disk $($originalDisk.Name) has been dettached from virtual machine $($vmNameItem.name)"

                                Add-AzVMDataDisk -VM $vmToRestore -ManagedDiskId $restoreDisk.Id -Name $restoreDisk.Name -CreateOption Attach -Lun $originalDiskLun | Out-Null
                                $updateResult = Update-AzVM -ResourceGroupName $vmToRestore.ResourceGroupName -VM $vmToRestore
                                if (!$updateResult.IsSuccessStatusCode) {
                                    $restoreVmSuccess = $false
                                }
                            }
                            else {
                                Write-Output "ERROR - Dettach original disk $($originalDisk.Name) from virtual machine $($vmNameItem.name) failed"
                                $restoreVmSuccess = $false
                            }                            
                        }

                        if ($updateResult.IsSuccessStatusCode) {
                            Write-Output "INFORMATION - New disk $($restoreDisk.Name) has been attached to virtual machine $($vmNameItem.name)"
                        }
                        else {
                            Write-Output "ERROR - Attach new disk $($restoreDisk.Name) to virtual machine $($vmNameItem.name) failed"
                            $restoreVmSuccess = $false
                        }
                    }
                    else {
                        Write-Output "ERROR - Create new disk $($restoreDisk.Name) from snapshot $($snapshotToRestore.Name) failed"
                        $restoreVmSuccess = $false
                    }
                }

                $snapshotToRestore = $null
                $originalDiskLun = $null
            }
        }
        #endregion Restore disk with new name

        #region Restore disk failed
        else {
            Write-Output "WARNING - Check if any snapshots exist (number of snapshots found $($snapshotForRestoreList.Count))"
            $restoreVmSuccess = $false
        }
        #endregion Restore disk failed
        
        if ($restoreVmSuccess) {
            $vmStartResult = $vmToRestore | Start-AzVM
            if (($vmStartResult.Status -like "Succeeded")) {
                Write-Output "INFORMATION - Virtual machine $($vmNameItem.name) is powered on"
            }
            else {
                Write-Output "WARNING - Power on virtual machine $($vmNameItem.name) failed"
            }
        }
        else {
            Write-Output "ERROR - Restore virtual machine $($vmNameItem.name) failed"
            $restoreVmSuccess = $true
        }
    }
    
    $vmToRestoreDiskList.Clear()
}
Write-Output "INFORMATION - Runbook has finished"
# end of script