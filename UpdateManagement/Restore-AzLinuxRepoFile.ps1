[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]
    $PatchWaveTag
)

#region Base Session Config
$VerbosePreference                        = 'SilentlyContinue'
$null = Get-Module    -Name 'AzureRM.*' | Remove-Module -Force
$null = Import-Module -Name 'Az.Accounts', 'Az.Automation', 'Az.Storage', 'Az.Compute', 'Az.ResourceGraph'
Trap { $_ | fl *; Write-Error -Message $_.Exception.Message -EA 1 }
$ConfirmPreference                        = 'None'
$VerbosePreference                        = 'Continue'
$ErrorActionPreference                    = 'Continue'
$PSDefaultParameterValues.'*:ErrorAction' = 'Stop'
#endregion Base Session Config

#region Account connection
Write-Output "INFORMATION - Authorizing automation account"
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

#region Get list of virtual machines
# Create Azure Resource Graph query and get desired VMs
# To be sure AA gets only virtual machine that manages, workspace ID used by OMS extenstion installed on VM is used in query
Write-Output "INFORMATION - Creating Azure Resource Graph query"
$aaLawId = Get-AutomationVariable -Name 'AaLawId'
$argQuery = [System.Text.StringBuilder]::new()
$argQuery.Append("resources ") | Out-Null
$argQuery.Append("| where type =~ 'microsoft.compute/virtualMachines' ") | Out-Null
$argQuery.Append("| where properties.storageProfile.osDisk.osType has 'Linux' ") | Out-Null
$argQuery.Append("and tags.PatchWave has '$PatchWaveTag' ") | Out-Null
$argQuery.Append("and properties.extended.instanceView.powerState.displayStatus has 'VM running' ") | Out-Null
$argQuery.Append("| extend ") | Out-Null
$argQuery.Append("JoinID = toupper(id) ") | Out-Null
$argQuery.Append("| join kind=leftouter( ") | Out-Null
$argQuery.Append("resources ") | Out-Null
$argQuery.Append("| where type == 'microsoft.compute/virtualmachines/extensions' ") | Out-Null
$argQuery.Append("| where name has 'OMS' ") | Out-Null
$argQuery.Append("| extend ") | Out-Null
$argQuery.Append("VMId = toupper(substring(id, 0, indexof(id, '/extensions'))),") | Out-Null
$argQuery.Append("workspaceId = properties.settings.workspaceId ") | Out-Null
$argQuery.Append(") on ") | Out-Null
$argQuery.Append("$") | Out-Null
$argQuery.Append("left.JoinID == ") | Out-Null
$argQuery.Append("$") | Out-Null
$argQuery.Append("right.VMId ") | Out-Null
$argQuery.Append("| where workspaceId == '") | Out-Null
$argQuery.Append($aaLawId) | Out-Null
$argQuery.Append("' ") | Out-Null # Change workspace ID to AA variable
$argQuery.Append("| project name, resourceGroup, subscriptionId, patchWave=tags.PatchWave") | Out-Null

Write-Output "AZURE RESOURCE GRAPH query begin"
$argQuery.ToString()
Write-Output "AZURE RESOURCE GRAPH query end"
$linuxQueryResult = Search-AzGraph -Query $argQuery.ToString() -First 1000
#FIXME If no virtual machine is found, whole patch management should be stopped to avoid running update with all repositories
#endregion Get list of virtual machines

#region Download shell script
Write-Output "INFORMATION - Downloading shell script content"
$AaStorageAccount = Get-AutomationVariable -Name AAccStorageAccount
$AaStorageAccountKeyOne = Get-AutomationVariable -Name AAccStorageAccountKeyOne
$saContext = New-AzStorageContext -StorageAccountName $AaStorageAccount -StorageAccountKey $AaStorageAccountKeyOne
$null = Get-AzStorageBlobContent -Blob "linuxrestorerepofiles.sh" -Container "invokecommands" -Context $saContext -Destination "C:\Temp"
#endregion Download shell script

#region Invoke shell script on virtual machine
Write-Output "INFORMATION - Process to restore repository files in progress"
foreach ($linuxItem in $linuxQueryResult) {
    $linuxObject = [PSCustomObject]@{
        VmName = $linuxItem.name
        RgName = $linuxItem.resourceGroup
        SubscriptionId = $linuxItem.subscriptionId
        PatchWave = $linuxItem.patchWave
    }

    Set-AzContext -SubscriptionId $linuxObject.SubscriptionId

    Write-Output "INFORMATION - Invoking command on VmName:$($linuxObject.VmName) PatchWave:$($linuxObject.PatchWave)"
	$invokeResult = Invoke-AzVMRunCommand -ResourceGroupName $linuxObject.RgName -Name $linuxObject.VmName -CommandId 'RunShellScript' -ScriptPath "C:\Temp\linuxrestorerepofiles.sh"

	if ($invokeResult.Status -like "Succeeded") {
		Write-Output "INFORMATION - Invoke command has succeeded"
	}
	else {
		Write-Output "ERROR - Invoke command has failed"
	}
}
#region Invoke shell script on virtual machine
Write-Output "INFORMATION - Process to restore repository files has completed"