#parameters section

param (

    # generator parameters are mandatory
    [Parameter(Mandatory = $true)][string]$outputPath, # output location where the result generated csv file needs to be placed
    [Parameter(Mandatory = $true)][string]$salesReportTemplatePath, # path to the generated SHIP salesreport file 
    [Parameter(Mandatory = $true)][int]$returnUnitsFirstItem, # number of units to be returned for the 1st item in the order

    # generator parameters are optional - if generated SHIP salesreport file contains more than one item 
    [Parameter(Mandatory = $true)][int]$returnUnitsSecondItem,
    [Parameter(Mandatory = $true)][int]$returnUnitsThirdItem
)


try {

    # write result parameters into a console for logging  
Write-Host @"
Generator parameters:
outputPath: $outputPath
salesReportTemplatePath: $salesReportTemplatePath
"@

if ($returnUnitsFirstItem -gt 0) {
    Write-Host @"
unitsToBeReturned for the 1st item: $returnUnitsFirstItem
"@
}

if ($returnUnitsSecondItem -gt 0) {
    Write-Host @"
unitsToBeReturned for the 2nd item: $returnUnitsSecondItem
"@
}

if ($returnUnitsThirdItem -gt 0) {
    Write-Host @"
unitsToBeReturned for the 3rd item: $returnUnitsThirdItem
"@
}
    $originalCSV = Import-Csv -Path $salesReportTemplatePath -Delimiter ";"

    $output = @()
    function GetReturnnewRecord($newRecord, $unitsToBeReturned) {
    if ($unitsToBeReturned -gt 0) {
        $newRecord.TYPE = "RETURN"
        $newRecord.QUANTITY = $unitsToBeReturned.ToString()
        return $newRecord
    }
    return $null
}

    $newRecord1 = GetReturnnewRecord $originalCSV[0] $returnUnitsFirstItem
    if ($newRecord1) { $output += $newRecord1 }

    $newRecord2 = GetReturnnewRecord $originalCSV[1] $returnUnitsSecondItem
    if ($newRecord2) { $output += $newRecord2 }

    $newRecord3 = GetReturnnewRecord $originalCSV[2] $returnUnitsThirdItem
    if ($newRecord3) { $output += $newRecord3 }


    $timestamp = Get-Date -Format "yyyy-MM-ddTHHmmss"
    $channelSign = $originalCSV[0].CHANNEL_SIGN
    $outputFileName = "salesreport_${channelSign}_${timestamp}.csv"


    $outputFileFullPath = Join-Path $outputPath $outputFileName
    
    Write-Host "Total newRecords to export: $($output.Count)"
    $output | Export-Csv -Path $outputFileFullPath -NoTypeInformation -Delimiter ";"
    Write-Host "Salesreport file $outputFileFullPath generated successfully"
    
} catch [Exception] {

    $errorMessage = "FAILED: $_"
    Write-Error $errorMessage
}
