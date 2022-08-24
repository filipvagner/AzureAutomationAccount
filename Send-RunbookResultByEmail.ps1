# It is not possible to use parameter sets in automation account (as of November 2021)
# Default set for both options below parameter are
# ToEmailAddress, FromDisplayName (optional), FromEmailAddress (optional), EmailSubject
#
# To send messages from Log Analytics Workspace, use as parameter set parameters below
# LogAnalyticsWorskspace, RunbookName, JobId, AllMessages (optional), WaitMinutes (optional)
#
# To send custom body, use as parameter set parameters below
# CustomEmailBody, EmailBody

Param(
    [Parameter(Mandatory=$true)]
    [string] $ToEmailAddress,
    [Parameter(Mandatory=$false)]
    [string] $FromDisplayName = "Default name",
    [Parameter(Mandatory=$false)]
    [string] $FromEmailAddress = "name@domain",
    [Parameter(Mandatory=$true)]
    [string] $EmailSubject,
    [Parameter(Mandatory=$false)]
    [bool] $LogAnalyticsWorskspace,
    [Parameter(Mandatory=$false)]
    [string] $RunbookName,
    [Parameter(Mandatory=$false)]
    [string] $JobId,
    [Parameter(Mandatory=$false)]
    [string] $RunbookInputList,
    [Parameter(Mandatory=$false)]
    [bool] $AllMessages = $false,
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 20)]
    [int] $WaitMinutes = 5,
    [Parameter(Mandatory=$false)]
    [bool] $CustomEmailBody,
    [Parameter(Mandatory=$false)]
    [string] $EmailBody
)

#region Module operation
$null = Get-Module -Name 'AzureRM.*' | Remove-Module -Force
$null = Import-Module -Name "Az.OperationalInsights", "Az.Accounts", "Az.Automation", "Az.KeyVault"
#endregion Module operation

#region Account connection
$Conn = Get-AutomationConnection -Name AzureRunAsConnection
$ConnectAzAccountParams = @{
  ServicePrincipal      = $true
  Tenant                = $Conn.TenantID
  ApplicationId         = $Conn.ApplicationID
  CertificateThumbprint = $Conn.CertificateThumbprint
}
$null = Connect-AzAccount @ConnectAzAccountParams
$null = Set-AzContext -Subscription (Get-AutomationVariable -Name "AAccSubscription")
#endregion Account connection

#region Process Information
if ($LogAnalyticsWorskspace) {
    #region Get parent's runbook log
    Write-Output "INFORMATION - Gathering logs from Log Analytics Workspace"
    $lawQuery = "
    AzureDiagnostics
    | where JobId_g == ""$jobId""
    | where Category == 'JobStreams'
    | sort by TimeGenerated asc
    | project TimeGenerated, RunbookName_s, ResultDescription
    "
    $logWorkspaceId = Get-AutomationVariable -Name "AaLawId"
    $waitCounter = 0
    do {
        Start-Sleep -Seconds 60
        $waitCounter++
        if ($waitCounter -eq $WaitMinutes) {
            Write-Output "ERROR - Counter reached limit"
            Write-Output "INFOMRATION - Check Log Analytics Workspace if contains parents job log messages"
            exit
        }
        $lawResults = Invoke-AzOperationalInsightsQuery -WorkspaceId $logWorkspaceId -Query $lawQuery
    } while ([System.String]::IsNullOrEmpty($lawResults.Results.ResultDescription))

    $lawLogList = New-Object -TypeName "System.Collections.ArrayList"
    foreach ($lawItem in $lawResults.Results) {
        $lawObject = [PSCustomObject]@{
            TimeGenerated = $lawItem.TimeGenerated
            RunbookName = $lawItem.RunbookName_s
            Message = $lawItem.ResultDescription
        }
        $lawLogList.Add($lawObject)
    }

    if ($lawLogList.Count -eq 0) {
        Write-Output "ERROR - No messages yield from log analytics workspace"
        Write-Output "INFORMATION - Runbook has stopped"
        exit
    }
    #endregion Get parent's runbook log

    #region Parse email body
    Write-Output "INFORMATION - Parsing email body and properties"
    $warningCounter = 0
    $errorCounter = 0
    foreach ($logItem in $lawLogList) {
        if ($logItem.Message.StartsWith("WARNING")) {
            $warningCounter++
        }
        elseif ($logItem.Message.StartsWith("ERROR")) {
            $errorCounter++
        }
    }

    $emailBodyString = [System.Text.StringBuilder]::new()
    $emailBodyString.Append("Runbook <strong>$RunbookName</strong> completed. To get details, check job's ID <strong>$JobId</strong>.<br>") | Out-Null
    $emailBodyString.Append("<br>") | Out-Null
    $emailBodyString.Append("Overview of parameters passed to runbook:") | Out-Null
    $emailBodyString.Append("<style>
    table {
    border-collapse: collapse;
    }

    td, th {
    border: 1px solid #dddddd;
    text-align: left;
    padding: 3px;
    }
    </style>") | Out-Null
    $emailBodyString.Append("
    <table>
    <tr>
        <th>Parameter Name</th>
        <th>Input Value</th>
    </tr>") | Out-Null
    foreach ($paramItem in $RunbookInputList.Split('|')) {
        $emailBodyString.Append("<tr>") | Out-Null
        $emailBodyString.Append("<td>") | Out-Null
        $emailBodyString.Append($paramItem.Split(':')[0]) | Out-Null
        $emailBodyString.Append("</td>") | Out-Null
        $emailBodyString.Append("<td>") | Out-Null
        $emailBodyString.Append($paramItem.Split(':')[1]) | Out-Null
        $emailBodyString.Append("</td>") | Out-Null
        $emailBodyString.Append("</tr>") | Out-Null
    }
    $emailBodyString.Append("</table>")
    $emailBodyString.Append("<br>") | Out-Null
    $emailBodyString.Append("Number of warning and error log messages found:") | Out-Null
    $emailBodyString.Append("<br>") | Out-Null
    $emailBodyString.Append("Warnings: <strong>$warningCounter</strong>") | Out-Null
    $emailBodyString.Append("<br>") | Out-Null
    $emailBodyString.Append("Errors: <strong>$errorCounter</strong>") | Out-Null
    $emailBodyString.Append("<br>") | Out-Null
    $emailBodyString.Append("<br>") | Out-Null

    if ($AllMessages) {
        $emailBodyString.Append("<strong>Detail of all log messages below:</strong>") | Out-Null
        $emailBodyString.Append("<br>") | Out-Null

        foreach ($logItem in $lawLogList) {
        if (($logItem.Message.StartsWith("INFORMATION")) -or ($logItem.Message.StartsWith("WARNING")) -or ($logItem.Message.StartsWith("ERROR"))) {
                $emailBodyString.Append($logItem.TimeGenerated + " - " + $logItem.Message) | Out-Null
                $emailBodyString.Append("<br>") | Out-Null
            }
        }
    } else {
        $emailBodyString.Append("<strong>Detail of warning and error log messages below:</strong>") | Out-Null
        $emailBodyString.Append("<br>") | Out-Null

        foreach ($logItem in $lawLogList) {
        if (($logItem.Message.StartsWith("WARNING")) -or ($logItem.Message.StartsWith("ERROR"))) {
                $emailBodyString.Append($logItem.TimeGenerated + " - " + $logItem.Message) | Out-Null
                $emailBodyString.Append("<br>") | Out-Null
            }
        }
    }

    $emailBodyString.Append("<br>") | Out-Null
    $emailBodyString.Append("For more information please contact <a href = ""mailto: name@domain"">Name of team</a>.") | Out-Null

    $emailBodySg = $emailBodyString.ToString()
    #endregion Parse email body
}
elseif ($CustomEmailBody) {
    $emailBodySg = $EmailBody
}
else {
    throw "You must specify parameter 'LogAnalyticsWorskspace' or 'CustomEmailBody', automation account does not support parameter sets"
}
#endregion Process Information

#region Parse recipient email addresses
$toRecepients = @()
foreach ($emailItem in $ToEmailAddress.Split(',')) {
    $toRecepients = $toRecepients + @{"email" = $emailItem}
}
#endregion Parse recipient email addresses


#region Send email
Write-Output "INFORMATION - Sending email"
$vaultName = Get-AutomationVariable -Name "AaKeyVault"
$sendGridApiKeySec = (Get-AzKeyVaultSecret -VaultName $vaultName -Name "SendGridRubookResultEmail").SecretValue
$sendGridApiKeyToBstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sendGridApiKeySec)
$sendGridApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($sendGridApiKeyToBstr)
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer " + $sendGridApiKey)
$headers.Add("Content-Type", "application/json")

$sendGridHeader = @{
    "authorization" = "Bearer $sendGridApiKey"
}

$sendGridBody = @{
    "personalizations" = @(
        @{
            "to"      = $toRecepients
            "subject" = $EmailSubject
        }
    )
    "content" = @(
        @{
            "type"  = "text/html"
            "value" = $emailBodySg
        }
    )
    "from" = @{
        "email" = $FromEmailAddress
        "name"  = $FromDisplayName
    }
}
$sendGridBodyJson = $sendGridBody | ConvertTo-Json -Depth 4

$sendGridParameters = @{
    Method      = "POST"
    Uri         = "https://api.sendgrid.com/v3/mail/send"
    Headers     = $sendGridHeader
    ContentType = "application/json"
    Body        = $sendGridBodyJson
}

Invoke-RestMethod @sendGridParameters
#endregion Send email