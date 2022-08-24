###########################################################
# AUTHOR  : Filip Vagner
# EMAIL   : filip.vagner@hotmail.com
# DATE    : 20-03-2022 (dd-mm-yyyy)
# COMMENT : This script based on reference list and list for each automation account with results compared to reference list,
#           creates table with visual results.
#           It is expecting list (CSV file) with following headers name,subscription,resourcegroup
#           NOTE: Before to run following script, all steps below should be completed
#           1. Prepare reference list (default modules + custom modules) > Get-ModuleReferenceList
#           2. Get current modules installed on each AA > Get-AzAaModule
#           3. Generate list for each AA with with compare to reference list > Compare-AzAaModule
###########################################################

function Get-AaModuleStatus ($DataFilePath) {
    $aaModulePath = $DataFilePath
    $aaModuleData = Import-Csv -Path $aaModulePath
    $aaModuleDataList = New-Object -TypeName "System.Collections.ArrayList"
    $aaModuleData | ForEach-Object {
        $aaModuleDataObj = [PSCustomObject]@{
            Name = $_.Name
            ModuleVersionStatus = $_.ModuleVersionStatus
            VersionAa = $_.VersionAa
            VersionRequired = $_.VersionRequired
            Global = $_.Global
        }
        $aaModuleDataList.Add($aaModuleDataObj) | Out-Null
    }

    return $aaModuleDataList
}

# Other Variables
$currentDateTime = Get-Date -Format "yyyy-MM-dd hh:mm:ss" -AsUTC

# Module Data
$refModulePath = "<path>\AaModuleManagement\module_ref_list.csv"
$refModuleData = Import-Csv -Path $refModulePath
$refModuleDataList = New-Object -TypeName "System.Collections.ArrayList"
$refModuleData | ForEach-Object {
    $refModuleDataObj = [PSCustomObject]@{
        Name = $_.Name
        Version = $_.Version
        Global = $_.Global
    }
    $refModuleDataList.Add($refModuleDataObj) | Out-Null
}
$aaDataFiles = Get-ChildItem -Path "<path>\AaModuleManagement\aa-compared-modules"

# Arrray definition
$rowNum = $refModuleDataList.Count+1
$colNum = ($aaDataFiles.Count-1)+3 # -1 Because there is folder inside
$moduleStatArr = New-Object 'object[,]' $rowNum, $colNum
$moduleStatArr[0, 0] = "Module Name"
$moduleStatArr[0, 1] = "Version"
$moduleStatArr[0, 2] = "Global"
$colCounter = 3
$aaDataFiles | Where-Object {$_.Name.EndsWith(".csv")} | Select-Object Name | ForEach-Object {$moduleStatArr[0, $colCounter] = $_.Name.Split('.')[0]; $colCounter++}

# Loading data into array
for ($i = 0; $i -lt $refModuleDataList.Count; $i++) {
    $colCounter = 0
    $rowCoutner = $i + 1
    $moduleStatArr[$rowCoutner,$colCounter] = $refModuleDataList[$i].Name
    $colCounter++
    $moduleStatArr[$rowCoutner,$colCounter] = $refModuleDataList[$i].Version
    $colCounter++
    $moduleStatArr[$rowCoutner,$colCounter] = $refModuleDataList[$i].Global
    $colCounter++
    
    foreach ($aaDataItem in $aaDataFiles | Where-Object {$_.Name.EndsWith(".csv")}) {
        $aaModuleList = Get-AaModuleStatus -DataFilePath $aaDataItem.FullName
        $moduleStatArr[$rowCoutner,$colCounter] = $aaModuleList[$i].ModuleVersionStatus
        $colCounter++
    }
}

# Table creation
$tableBodyString = [System.Text.StringBuilder]::new()
$tableBodyString.Append("<style>
table {
    border-collapse: collapse;
}

td, th {
    border: 1px solid #dddddd;
    text-align: left;
    padding: 3px;
}
</style>") | Out-Null
$tableBodyString.Append("<table>") | Out-Null
$tableBodyString.Append("<tr>") | Out-Null

for ($n = 0; $n -lt $colNum; $n++) {
    $cell = $moduleStatArr[0,$n]
    $tableBodyString.Append("<th>") | Out-Null
    $tableBodyString.Append($cell) | Out-Null
    $tableBodyString.Append("</th>") | Out-Null
}
$tableBodyString.Append("</tr>") | Out-Null

for ($j = 0; $j -lt $rowNum; $j++) {
    $colCounter = 0
    $rowCoutner = $j + 1
    $tableBodyString.Append("<tr>") | Out-Null
    do {
        $cell = $moduleStatArr[$rowCoutner,$colCounter]
        if ($cell -like "Equal") {
            $tableBodyString.Append("<td style=""background-color:#90EE90"">") | Out-Null
            $tableBodyString.Append($cell) | Out-Null
            $tableBodyString.Append("</td>") | Out-Null
        }
        elseif ($cell -like "RequiredHigher") {
            $tableBodyString.Append("<td style=""background-color:#F08080"">") | Out-Null
            $tableBodyString.Append($cell) | Out-Null
            $tableBodyString.Append("</td>") | Out-Null
        }
        elseif ($cell -like "Missing") {
            $tableBodyString.Append("<td style=""background-color:#FFA07A"">") | Out-Null
            $tableBodyString.Append($cell) | Out-Null
            $tableBodyString.Append("</td>") | Out-Null
        }
        elseif ($cell -like "RequiredLower") {
            $tableBodyString.Append("<td style=""background-color:#D3D3D3"">") | Out-Null
            $tableBodyString.Append($cell) | Out-Null
            $tableBodyString.Append("</td>") | Out-Null
        }
        else {
            $tableBodyString.Append("<td>") | Out-Null
            $tableBodyString.Append($cell) | Out-Null
            $tableBodyString.Append("</td>") | Out-Null
        }
        
        $colCounter++
    } while ($colCounter -lt $colNum)
    $tableBodyString.Append("</tr>") | Out-Null
}
$tableBodyString.Append("</table>") | Out-Null
$tableBodyString.Append("<br>") | Out-Null
$tableBodyString.Append("Content generated (UTC): ") | Out-Null
$tableBodyString.Append($currentDateTime) | Out-Null
$tableBodyString.Append("<br>") | Out-Null
$tableBodyString.Append("Cell explanation:") | Out-Null
$tableBodyString.Append("<ul>") | Out-Null
$tableBodyString.AppendLine("<li>Equal - Installed module is same version as required</li>") | Out-Null
$tableBodyString.Append("<li>Missing - Module is not installed</li>") | Out-Null
$tableBodyString.Append("<li>RequiredHigher - Installed module is lower version than required (upgrade needed)</li>") | Out-Null
$tableBodyString.Append("<li>RequiredLower - Installed module is higher version than required</li>") | Out-Null
$tableBodyString.Append("</ul>") | Out-Null

$tableBodyString.ToString() | Out-File -FilePath "<path>\AaModuleManagement\module_status_result.html" 