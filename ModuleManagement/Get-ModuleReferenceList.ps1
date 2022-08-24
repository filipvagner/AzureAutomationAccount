###########################################################
# AUTHOR  : Filip Vagner
# EMAIL   : filip.vagner@hotmail.com
# DATE    : 17-03-2022 (dd-mm-yyyy)
# COMMENT : This script takes two lists,list of default modules and list of custom  modules.
#           Default module list, is list of modules that appear on automation account after it's creation 
#           (so should be checked regulary, can be obtained on https://docs.microsoft.com/en-us/azure/automation/shared-resources/modules or create AA) and should be used as reference list.
#           Custom list, is unique list of all modules that has been uploaded to automation account after it's creation and are not part of default list 
#           (if uploaded module with name is same as default, then default module is overwritten).
#           It is expecting each list (CSV file) with following headers name,version
#
###########################################################

$moduleDefaultDataPath = "<path>\AaModuleManagement\module_default.csv"
$moduleDefaultData = Import-Csv -Path $moduleDefaultDataPath
$moduleCustomDataPath = "<path>\AaModuleManagement\module_custom.csv"
$moduleCustomData = Import-Csv -Path $moduleCustomDataPath
$refModuleListPath = "<path>\AaModuleManagement\module_ref_list.csv"

# Default Module Data
$moduleDefaultDataList = New-Object -TypeName "System.Collections.ArrayList"
$moduleDefaultData | ForEach-Object {
    $moduleDefaultDataObj = [PSCustomObject]@{
        Name = $_.name
        Version = $_.version
        Global = $true
    }
    $moduleDefaultDataList.Add($moduleDefaultDataObj) | Out-Null
}

# Custom Module Data
$moduleCustomDataList = New-Object -TypeName "System.Collections.ArrayList"
$moduleCustomData | ForEach-Object {
    $moduleCustomDataObj = [PSCustomObject]@{
        Name = $_.name
        Version = $_.version
        Global = $false
    }
    $moduleCustomDataList.Add($moduleCustomDataObj) | Out-Null
}

# Compare if module name is on list
$moduleRefDataList = New-Object -TypeName "System.Collections.ArrayList"
$moduleSupDataList = New-Object -TypeName "System.Collections.ArrayList"

if ($moduleDefaultDataList.Count -gt $moduleCustomDataList.Count) {
    $moduleRefDataList = $moduleDefaultDataList
    $moduleSupDataList = $moduleCustomDataList
} 
elseif ($moduleDefaultDataList.Count -lt $moduleCustomDataList.Count){
    $moduleRefDataList = $moduleCustomDataList
    $moduleSupDataList = $moduleDefaultDataList
}
else {
    $moduleRefDataList = $moduleDefaultDataList
    $moduleSupDataList = $moduleCustomDataList

    $moduleDataObj = [PSCustomObject]@{
        Name = "ThisIsFake"
        Version = "1.0"
        Global = $false
    }
    $moduleRefDataList.Add($moduleDataObj) | Out-Null
}

for ($i = 0; $i -lt $moduleRefDataList.Count; $i++) {
    for ($j = 0; $j -lt $moduleSupDataList.Count; $j++) {
        if ($moduleSupDataList[$j].Name -like $moduleRefDataList[$i].Name) {
            if ($moduleRefDataList[$i].Version -lt $moduleSupDataList[$j].Version) {
                $moduleRefDataList[$i].Version = $moduleSupDataList[$j].Version
                $moduleRefDataList[$i].Global = $false
            }
        }
    }
}

foreach ($moduleItem in $moduleSupDataList) {
    if (!$moduleRefDataList.Name.Contains($moduleItem.Name)) {
        $moduleRefDataList.Add($moduleItem) | Out-Null
    }
}

$moduleRefDataList = $moduleRefDataList | Sort-Object -Property Name
$moduleRefDataList | ConvertTo-Csv | Out-File $refModuleListPath