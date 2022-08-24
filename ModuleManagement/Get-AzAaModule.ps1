###########################################################
# AUTHOR  : Filip Vagner
# EMAIL   : filip.vagner@hotmail.com
# DATE    : 17-03-2022 (dd-mm-yyyy)
# COMMENT : This script calls Azure API and gets list of modules installed on automation account.
#           It is expecting list (CSV file) with following headers name,subscription,resourcegroup
#
###########################################################

$aaDataPath = "<path>\AaModuleManagement\aa_list.csv"
$aaData = Import-Csv -Path $aaDataPath
$aaDataList = New-Object -TypeName "System.Collections.ArrayList"
$aaData | ForEach-Object {
    $aaDataObj = [PSCustomObject]@{
        Name = $_.name
        Subscription = $_.subscription
        Resourcegroup = $_.resourcegroup
    }
    $aaDataList.Add($aaDataObj) | Out-Null
}

$azToken = (Get-AzAccessToken).Token
$aaRequestHeader = @{
    "authorization" = "Bearer $azToken"
}

foreach ($aaItem in $aaDataList) {
    $targetUri = "https://management.azure.com/subscriptions/" + $aaItem.Subscription +"/resourceGroups/" + $aaItem.Resourcegroup + "/providers/Microsoft.Automation/automationAccounts/" +  $aaItem.Name + "/modules?api-version=2015-10-31"
    $aaRequest = @{
        Method      = "GET"
        Uri         = $targetUri
        Headers     = $aaRequestHeader
    }
    $aaModuleObjectList = New-Object -TypeName "System.Collections.ArrayList"

    do {
        $aaRequestContent = Invoke-WebRequest @aaRequest
        $aaModuleObject = ConvertFrom-Json -InputObject $aaRequestContent.Content
        
        foreach ($aaModule in $aaModuleObject.value) {
            $aaModulgObj = [PSCustomOBject]@{
                Name = $aaModule.name
                Version = 
                    if ($null -eq $aaModule.properties.version){
                    "0.0"  
                    }
                    else {
                        $aaModule.properties.version
                    }
                Global = $aaModule.properties.isGlobal
            }
            $aaModuleObjectList.Add($aaModulgObj) | Out-Null
        }

        # By default only 100 records is returned, response returns link for another call
        $aaRequest.Uri = $aaModuleObject.nextLink
    } until ($null -eq $aaModuleObject.nextLink)

    $aaModuleObjectList = $aaModuleObjectList | Sort-Object -Property Name
    $outputPath = "<path>\AaModuleManagement\aa-current-modules\" + $aaItem.Name + ".csv"
    $aaModuleObjectList | ConvertTo-Csv | Out-File $outputPath
}