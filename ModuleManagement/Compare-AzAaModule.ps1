###########################################################
# AUTHOR  : Filip Vagner
# EMAIL   : filip.vagner@hotmail.com
# DATE    : 17-03-2022 (dd-mm-yyyy)
# COMMENT : This script takes two lists. List of modules that is used as reference and list of modules on desired automation account.
#           These two lists are compared and status is stored in property ModuleVersionStatus.
#           This property has following values:
#           Eaqual          = AA module has same version as module on reference list
#           RequiredHigher  = AA module has lower version than on reference list (should be upgraded)
#           RequiredLower   = AA module has higher version than on reference list (can be downgraded or reference list should be updated)
#           Missing         = AA module is missing
#           Error           = Unexpected result when comparing versions
#
###########################################################

$refModulePath = "<path>\AaModuleManagement\module_ref_list.csv"
$refModuleData = Import-Csv -Path $refModulePath
# Reference Module Data
$refModuleDataList = New-Object -TypeName "System.Collections.ArrayList"
$refModuleData | ForEach-Object {
    $refModuleDataObj = [PSCustomObject]@{
        Name = $_.Name
        Version = $_.Version
        Global = $_.Global
    }
    $refModuleDataList.Add($refModuleDataObj) | Out-Null
}

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

foreach ($aaItem in $aaDataList) {
    $aaModuleDataPath = "<path>\AaModuleManagement\aa-current-modules\" + $aaItem.Name + ".csv"
    $aaModuleData = Import-Csv -Path $aaModuleDataPath

    # AA Module Data
    $aaModuleDataList = New-Object -TypeName "System.Collections.ArrayList"
    $aaModuleData | ForEach-Object {
        $aaModuleDataObj = [PSCustomObject]@{
            Name = $_.Name
            Version = $_.Version
            Global = $_.Global
        }
        $aaModuleDataList.Add($aaModuleDataObj) | Out-Null
    }

    $moduleFinalDataList = New-Object -TypeName "System.Collections.ArrayList"
    for ($i = 0; $i -lt $refModuleDataList.Count; $i++) {
        if ($aaModuleDataList.Name.Contains($refModuleDataList[$i].Name)) {
            [version]$refModuleVersion = $refModuleDataList[$i].Version
            [version]$aaModuleVersion = ($aaModuleDataList | Where-Object {$_.Name -like $refModuleDataList[$i].Name}).Version
            
            if ($refModuleVersion -gt $aaModuleVersion) {
                $versionStatus = "RequiredHigher"
            }
            elseif ($refModuleVersion -lt $aaModuleVersion) {
                $versionStatus = "RequiredLower"
            }
            elseif ($refModuleVersion -eq $aaModuleVersion) {
                $versionStatus = "Equal"
            }
            else {
                $versionStatus = "Error"
            }

            $aaNewDataObj = [PSCustomObject]@{
                Name = $refModuleDataList[$i].Name
                ModuleVersionStatus = $versionStatus
                VersionAa = $aaModuleVersion
                VersionRequired = $refModuleVersion
                Global = $refModuleDataList[$i].Global
            }
            $moduleFinalDataList.Add($aaNewDataObj) | Out-Null
        }
        else {
            $aaNewDataObj = [PSCustomObject]@{
                Name = $refModuleDataList[$i].Name
                ModuleVersionStatus = "Missing"
                VersionAa = "Missing"
                VersionRequired = $refModuleDataList[$i].Version
                Global = $refModuleDataList[$i].Global
            }
            $moduleFinalDataList.Add($aaNewDataObj) | Out-Null
        }
    }

    $moduleFinalDataList = $moduleFinalDataList | Sort-Object -Property Name
    # $outputPath = "<path>\AaModuleManagement\aa-compared-modules\" + $aaItem.Name + ".txt"
    # $moduleFinalDataList | Format-Table | Out-File $outputPath
    $outputPath = "<path>\AaModuleManagement\aa-compared-modules\" + $aaItem.Name + ".csv"
    $moduleFinalDataList | ConvertTo-Csv | Out-File $outputPath
}