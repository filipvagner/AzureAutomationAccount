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

#region Functions
function Get-PreviousWeekDate ($DayToStart, $DaysBackward) {
    
    $weekDates = New-Object -TypeName "System.Collections.ArrayList"
    $tCulture = [System.Globalization.CultureInfo]::InvariantCulture
    
    $i = 0
    do {
        if ((Get-Date).AddDays(-$i).DayOfWeek -like $DayToStart) {
            for ($j = $DaysBackward; $j -gt 0; --$j) {
                $weekDates.Add((Get-Date).AddDays(((-$i) -$j) + 1).GetDateTimeFormats('d', $tCulture)[1]) | Out-Null
            }
            
            break
        }
        $i++
    } while ($true)

    return $weekDates
}
#endregion Functions

#region Chart config
# This region has been created to make mess below a bit more clear
# Data string pattern is replaced with real value before generating URL
$suChartConfig = "
{
    type: 'doughnut',
    data: {
      datasets: [
        {
          data: [%SUSP%, %SUFP%, %SUEP%],
          backgroundColor: [
            'rgb(71, 159, 64)',
            'rgb(231, 71, 71)',
            'rgb(232, 139, 58)'
          ],
          label: 'Dataset 1',
        },
      ],
      labels: ['Succeeded', 'Failed', 'Exceeded'],
    },
    options: {
      title: {
        display: true,
        text: 'Overall Software Update Status',
      },
      plugins: {
        datalabels: {
          color: 'white',
          formatter: (value) => {
              return value + '%';
            }
        }
      }
    },
  }
"

$vmChartConfig = "
{
    type: 'doughnut',
    data: {
      datasets: [
        {
          data: [%VMSP%, %VMFP%, %VMEP%],
          backgroundColor: [
            'rgb(71, 159, 64)',
            'rgb(231, 71, 71)',
            'rgb(232, 139, 58)'
          ],
          label: 'Dataset 1',
        },
      ],
      labels: ['Succeeded', 'Failed', 'Exceeded'],
    },
    options: {
      title: {
        display: true,
        text: 'Overall Machine Run Status',
      },
      plugins: {
        datalabels: {
          color: 'white',
          formatter: (value) => {
              return value + '%';
            }
        }
      }
    },
  }
"
#endregion Chart config

#region Variables
$scheduleRg = Get-AutomationVariable -Name 'AAccGroup'
$scheduleAa = Get-AutomationVariable -Name 'AAccName'
$emailRegion = ($scheduleAa.Split('-')[1]).ToUpper()
[string]$DayToStart = "Tuesday"
[int]$DaysBackward = 7
$weekDatesList = Get-PreviousWeekDate -DayToStart $DayToStart -DaysBackward $DaysBackward
$suRunList = Get-AzAutomationSoftwareUpdateRun -ResourceGroupName $scheduleRg -AutomationAccountName $scheduleAa -StartTime ([DateTimeOffset]::Parse($weekDatesList[0]))
[int]$suTotal = $suRunList.Count
[int]$suSucceeded = ($suRunList | Where-Object {$_.Status -like "Succeeded"}).Count
[int]$suSucceededPercent = (100 / ($suRunList.Count)) * ($suSucceeded)
[int]$suFailed = ($suRunList | Where-Object {$_.Status -like "Failed"}).Count
[int]$suFailedPercent = (100 / ($suRunList.Count)) * ($suFailed)
[int]$suExceeded = ($suRunList | Where-Object {$_.Status -like "MaintenanceWindowExceeded"}).Count
[int]$suExceededPercent = (100 / ($suRunList.Count)) * ($suExceeded)
[int]$vmTotal = 0
[int]$vmSucceeded = 0
[int]$vmFailed = 0
[int]$vmExceeded = 0
$suRunList | ForEach-Object {
  Get-AzAutomationSoftwareUpdateMachineRun -SoftwareUpdateRunId $_.RunId -ResourceGroupName $scheduleRg -AutomationAccountName $scheduleAa | ForEach-Object {
    switch ($_.Status) {
      "Succeeded" { $vmSucceeded++; break }
      "Failed" { $vmFailed++; break }
      "MaintenanceWindowExceeded" { $vmExceeded++; break }
      Default { $vmFailed++; break }
    }
  }
}
$vmTotal = $vmSucceeded + $vmFailed + $vmExceeded
[int]$vmSucceededPercent = (100 / $vmTotal) * $vmSucceeded
[int]$vmFailedPercent = (100 / $vmTotal) * $vmFailed
[int]$vmExceededPercent = (100 / $vmTotal) * $vmExceeded
$regexCheck = [System.Text.RegularExpressions.Regex]::new("\d\d\d\d-\d\d-\d\d", "ignoreCase")
$script:tCulture = [System.Globalization.CultureInfo]::InvariantCulture
$emailBodyString = [System.Text.StringBuilder]::new()
#endregion Variables

#region CSS Style
$emailBodyString.Append("<!DOCTYPE html>
<html>
<head>
<style>
.table {
  border-collapse: collapse;
  /*width: 100%;*/
}

.table td, .table th {
  border: 1px solid #ddd;
  padding: 8px;
}

.table tr:nth-child(even){background-color: #f2f2f2;}

.table tr:hover {background-color: #ddd;}

.table th {
  padding-top: 12px;
  padding-bottom: 12px;
  text-align: center;
  background-color: #008238;
  color: white;
}

.tdbold {
  font-weight: bold
}
.tdpadding {
  padding: 8px 23px!important;
}

</style>
</head>
") | Out-Null
#endregion CSS Style

$emailBodyString.Append("<body>
<table class='table'>
    <tr>
        <th colspan=""4""><h2>Update Management Report</h2></th>
    </tr>") | Out-Null
$emailBodyString.Append("<tr><td>Total number of deployment schedule</td>") | Out-Null
$emailBodyString.Append("<td class='tdbold tdpadding'>" + $suTotal + "</td>") | Out-Null
$emailBodyString.Append("<td>Total number of machines</td>") | Out-Null
$emailBodyString.Append("<td class='tdbold tdpadding'>" + $vmTotal + "</td></tr>") | Out-Null
$emailBodyString.Append("<tr><td>Successful number of deployment schedule</td>") | Out-Null
$emailBodyString.Append("<td class='tdbold tdpadding'>" + $suSucceeded + "</td>") | Out-Null
$emailBodyString.Append("<td>Successful number of machines</td>") | Out-Null
$emailBodyString.Append("<td class='tdbold tdpadding'>" + $vmSucceeded + "</td></tr>") | Out-Null
$emailBodyString.Append("<tr><td>Failed number of deployment schedule</td>") | Out-Null
$emailBodyString.Append("<td class='tdbold tdpadding'>" + $suFailed + "</td>") | Out-Null
$emailBodyString.Append("<td>Failed number of machines</td>") | Out-Null
$emailBodyString.Append("<td class='tdbold tdpadding'>" + $vmFailed + "</td></tr>") | Out-Null
$emailBodyString.Append("<tr><td>Exceeded maintenance number of deployment schedule</td>") | Out-Null
$emailBodyString.Append("<td class='tdbold tdpadding'>" + $suExceeded + "</td>") | Out-Null
$emailBodyString.Append("<td>Exceeded maintenance number of machines</td>") | Out-Null
$emailBodyString.Append("<td class='tdbold tdpadding'>" + $vmExceeded + "</td></tr></table>") | Out-Null
$emailBodyString.Append("<br>") | Out-Null
$suChartConfig = $suChartConfig.Replace("%SUSP%", $suSucceededPercent)
$suChartConfig = $suChartConfig.Replace("%SUFP%", $suFailedPercent)
$suChartConfig = $suChartConfig.Replace("%SUEP%", $suExceededPercent)
$suUrlToEncode = $suChartConfig
$suEncodedURL = [System.Web.HttpUtility]::UrlEncode($suUrlToEncode)
$suEncodedURL = "https://quickchart.io/chart?w=400&h=400&c=" + $suEncodedURL
[byte[]]$chartImg = (Invoke-WebRequest -Method GET -Uri $suEncodedURL -UseBasicParsing).Content
$chartImgB64 = [Convert]::ToBase64String($chartImg)
$emailBodyString.Append("<img src='data:image/png;base64, $($chartImgB64)' width=""400"" height=""400""/>") | Out-Null

$vmChartConfig = $vmChartConfig.Replace("%VMSP%", $vmSucceededPercent)
$vmChartConfig = $vmChartConfig.Replace("%VMFP%", $vmFailedPercent)
$vmChartConfig = $vmChartConfig.Replace("%VMEP%", $vmExceededPercent)
$vmUrlToEncode = $vmChartConfig
$vmEncodedURL = [System.Web.HttpUtility]::UrlEncode($vmUrlToEncode)
$vmEncodedURL = "https://quickchart.io/chart?w=400&h=400&c=" + $vmEncodedURL
[byte[]]$chartImg = (Invoke-WebRequest -Method GET -Uri $vmEncodedURL -UseBasicParsing).Content
$chartImgB64 = [Convert]::ToBase64String($chartImg)
$emailBodyString.Append("<img src='data:image/png;base64, $($chartImgB64)' width=""400"" height=""400""/>") | Out-Null

$emailBodyString.Append("<br>") | Out-Null
$emailBodyString.Append("
<h2>Details of deployment schedule</h2>
<p>Below is table with detailed information for each deployment schedule
    and list of servers in deployment schedules including name and status.
</p>") | Out-Null
$emailBodyString.Append("<hr><br>") | Out-Null
for ($n = 0; $n -lt $suRunList.Count; $n++) {
    if ($weekDatesList.Contains($regexCheck.Match($suRunList[$n].SoftwareUpdateConfigurationName).Value)) {
        $emailBodyString.Append("<table class='table'>") | Out-Null
        $emailBodyString.Append("<tr><th colspan=""4"" class='tdbold'>" + $suRunList[$n].SoftwareUpdateConfigurationName + "</th></tr>") | Out-Null
        $emailBodyString.Append("<tr><td colspan=""2"" class='tdbold'>Start time</td>") | Out-Null
        $emailBodyString.Append("<td colspan=""2"">" + $suRunList[$n].StartTime.LocalDateTime.GetDateTimeFormats('F', $tCulture) + "</td></tr>") | Out-Null
        $emailBodyString.Append("<tr><td colspan=""2"" class='tdbold'>End time</td>") | Out-Null
        try{
          $emailBodyString.Append("<td colspan=""2"">" + $suRunList[$n].EndTime.LocalDateTime.GetDateTimeFormats('F', $tCulture) + "</td></tr>") | Out-Null
        }catch{
          $emailBodyString.Append("<td colspan=""2"">N/A</td></tr>") | Out-Null
        }
        $emailBodyString.Append("<tr><td colspan=""2"" class='tdbold'>Overall status</td>") | Out-Null
        $emailBodyString.Append("<td colspan=""2"">" + $suRunList[$n].Status + "</td></tr>") | Out-Null
        $emailBodyString.Append("<td class='tdbold'>Machine Succeeded</td>") | Out-Null
        $emailBodyString.Append("<td class='tdbold'>" + ($suRunList[$n].ComputerCount - $suRunList[$n].FailedCount) + "</td>") | Out-Null
        $emailBodyString.Append("<td span class='tdbold'>Machine Failed</td>") | Out-Null
        $emailBodyString.Append("<td class='tdbold'>" + $suRunList[$n].FailedCount + "</td></tr>") | Out-Null
        $emailBodyString.Append("</table>") | Out-Null
        $emailBodyString.Append("<br>") | Out-Null
        
        if ($suRunList[$n].Status -like "Succeeded") {
          $emailBodyString.Append("<table class='table'>") | Out-Null
          $emailBodyString.Append("<tr><th colspan=""2""><strong>List of failed machines</strong></th></tr>") | Out-Null
          $emailBodyString.Append("<tr><td colspan=""2"">No failed machines in deployment schedule</td></tr>") | Out-Null
          $emailBodyString.Append("</table>") | Out-Null
          $emailBodyString.Append("<br>") | Out-Null
        }
        else {
          $emailBodyString.Append("<table class='table'>") | Out-Null
          $emailBodyString.Append("<tr><th colspan=""2""><strong>List of failed machines</strong></th></tr>") | Out-Null
          if($suRunList[$n].ComputerCount -eq 0){
            $emailBodyString.Append("<tr><td colspan=""2"">No machines in deployment schedule</td></tr>") | Out-Null
          
          }else{
            $emailBodyString.Append("<tr><td class='tdbold'>Machine name</td><td class='tdbold'>Status</td></tr>") | Out-Null
            $machineRunList = Get-AzAutomationSoftwareUpdateMachineRun -SoftwareUpdateRunId $suRunList[$n].RunId -ResourceGroupName $scheduleRg -AutomationAccountName $scheduleAa
            foreach ($machineItem in $machineRunList) {
              if (($machineItem.Status -like "Failed") -or ($machineItem.Status -like "FailedToStart") -or ($machineItem.Status -like "MaintenanceWindowExceeded")) {
                $emailBodyString.Append("<tr><td>" + $machineItem.TargetComputer.Split('/')[-1] + "</td>") | Out-Null
                $emailBodyString.Append("<td>" + $machineItem.Status + "</td></tr>") | Out-Null
              }
                
            }
          }
          
          $emailBodyString.Append("</table>") | Out-Null
          $emailBodyString.Append("<br>") | Out-Null
        }
        $emailBodyString.Append("<hr>") | Out-Null
        $emailBodyString.Append("<br>") | Out-Null
    }
}

$emailBodyString.Append("<p>
For more information please contact <a href = ""mailto: name@domain"">Name of team</a>.
</p>
</body>
</html>") | Out-Null

#region Send email
Write-Output "INFORMATION - Sending results by email"
$runbookResultEmailParams = @{
        ToEmailAddress = "name@domain"
        EmailSubject = "Update Schedule Report - $emailRegion"
        CustomEmailBody = $true
        EmailBody = $emailBodyString.ToString()
}

Start-AzAutomationRunbook `
    -AutomationAccountName $scheduleAa `
    -Name 'Send-RunbookResultByEmail' `
    -ResourceGroupName $scheduleRg `
    -Parameters $runbookResultEmailParams
#endregion Send email