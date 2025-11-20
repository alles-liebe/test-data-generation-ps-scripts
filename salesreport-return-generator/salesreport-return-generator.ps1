#parameters section

param (

    # generator parameters are mandatory
    [Parameter(Mandatory = $true)][string]$outputPath, # output location where the result generated csv file needs to be placed
    [Parameter(Mandatory = $true)][string]$orderTemplatePath, # path to the order.xml file 
    [Parameter(Mandatory = $true)][string]$stockLocationsPath, # path to Stock Locations names and IDs in JSON
    [Parameter(Mandatory = $true)][string]$sourceLocationsPath, # path to Source Locations names and IDs in JSON
    [Parameter(Mandatory = $true)][string]$selectedStockLocationName1, # LocationName for the 1st item selected in TeamCity the products are shipped from
    [Parameter(Mandatory = $true)][string]$selectedSourceLocationName1, # the original LocationName for the 1st item selected in TeamCity the products are picked from
    [Parameter(Mandatory = $true)][string]$selectedPackageType1, # the type of package selected in TeamCity for the 1st item: SHIP, CUST_CANCEL or NO_INVENTORY

    # generator parameters are optional - if order xml file contains more than one item 
    [Parameter(Mandatory = $false)][string]$selectedStockLocationName2, # LocationName for the 2nd item selected in TeamCity the products are shipped from
    [Parameter(Mandatory = $false)][string]$selectedSourceLocationName2, # the original LocationName for the 2nd item selected in TeamCity the products are picked from
    [Parameter(Mandatory = $false)][string]$selectedPackageType2, # the type of package selected in TeamCity for the 2nd item
    [Parameter(Mandatory = $false)][string]$selectedStockLocationName3, # LocationName for the 3rd item selected in TeamCity the products are shipped from
    [Parameter(Mandatory = $false)][string]$selectedSourceLocationName3, # the original LocationName for the 3rd item selected in TeamCity the products are picked from
    [Parameter(Mandatory = $false)][string]$selectedPackageType3 # the type of package selected in TeamCity for the 3rd item
    
)

try {

    # write result parameters into a console for logging 
    Write-Host @"
Generator parameters:
outputPath: $outputPath
orderTemplatePath: $orderTemplatePath
stockLocationsPath: $stockLocationsPath
sourceLocationsPath: $sourceLocationsPath
selectedStockLocationName1: $selectedStockLocationName1
selectedSourceLocationName1: $selectedSourceLocationName1
"@

    if (($selectedStockLocationName2) -and ($selectedSourceLocationName2)) {
        Write-Host @"
selectedStockLocationName2: $selectedStockLocationName2
selectedSourceLocationName2: $selectedSourceLocationName2
"@
        if (($selectedStockLocationName3) -and ($selectedSourceLocationName3)) {
            Write-Host @"
selectedStockLocationName3: $selectedStockLocationName3
selectedSourceLocationName3: $selectedSourceLocationName3
"@
        }
    }

    # read order template xml file
    $orderFileXml = [xml](Get-Content $orderTemplatePath)

    # get all elements from TradebyteOrder
    $order = $orderFileXml.ArrayOfTradebyteOrder.TradebyteOrder
    $orderData = $order.ORDER_DATA
    $orderDate = $orderData.ORDER_DATE.Split("T")[0]

    # load location configuration from JSON files
    $stockJson = Get-Content $stockLocationsPath -Raw | ConvertFrom-Json
    $sourceJson = Get-Content $sourceLocationsPath -Raw | ConvertFrom-Json

    # set salesreport csv file mask
    $channelSign = $orderData.CHANNEL_SIGN
    $timestamp = Get-Date -Format "yyyy-MM-ddTHHmmss"
    $outputFileName = "salesreport_${channelSign}_${timestamp}.csv"
    $outputFileFullPath = Join-Path (Split-Path $outputPath -Parent) $outputFileName

    # get all items from the order
    $items = @($order.ITEMS.ITEM)
    Write-Host "Number of items found: $($items.Count)"

    # log selected package type for the 1st item
    Write-Host "Type of the 1st package: '$selectedPackageType1'"

    # work on the mandatory parameters if order has one item or order should be shipped from the single location
    # find value where mandatory stock location name selected on TeamCity matches the ID in JSON
    $stockLocationId1 = ($stockJson | Where-Object { $_.name -eq $selectedStockLocationName1 }).id
    Write-Host "Found stock location ID for the 1st item: '$stockLocationId1'"

    # find value where mandatory source location name selected on TeamCity matches the ID in JSON
    $sourceLocationId1 = ($sourceJson | Where-Object { $_.name -eq $selectedSourceLocationName1 }).id
    Write-Host "Found source location ID for the 1st item: '$sourceLocationId1'"


    # build output CSV
    $output = @()

    # define the function to call for each item in the order
    function GetCSVOutputString($item, $stockId, $stockName, $sourceId, $sourceName, $selectedPackageType) {
        # random values for P_NR, A_NR, A_NR2, A_PROD_NR in the salesreport csv file for the item
        $pNr = Get-Random -Minimum 10000000 -Maximum 99999999
        $aNr = Get-Random -Minimum 10000000 -Maximum 99999999
        $idCode = "1Z" + (Get-Random -Minimum 100000000 -Maximum 999999999)

        # build a custom object for the item
        $outputString = [PSCustomObject]@{
            MESSAGE_DATE               = (Get-Date).ToString("yyyy-MM-dd")
            DATE_CREATED               = ($item.DATE_CREATED.Split("T")[0])
            CHANNEL_SIGN               = $orderData.CHANNEL_SIGN
            CHANNEL_ORDER_ID           = $orderData.CHANNEL_ID
            CHANNEL_ORDER_SHIPMENT_ID  = $orderData.CHANNEL_ID
            CHANNEL_ORDER_ITEM_ID      = $item.CHANNEL_ID
            CHANNEL_MESSAGE_ID         = ""
            TB_MESSAGE_ID              = ""
            TB_ORDER_ITEM_ID           = $item.TB_ID
            TB_ORDER_ID                = $orderData.TB_ID
            TYPE                       = $selectedPackageType
            QUANTITY                   = $item.QUANTITY
            P_NR                       = $pNr
            A_NR                       = $aNr
            A_NR2                      = $aNr
            A_PROD_NR                  = $item.SKU
            A_EAN                      = $item.EAN
            POS_ANR_CHANNEL            = $item.CHANNEL_SKU
            BILLING_TEXT               = $item.BILLING_TEXT
            ITEM_PRICE                 = $item.ITEM_PRICE.'#text'
            TRANSFER_PRICE             = $item.TRANSFER_PRICE.'#text'
            POS_CONDITION              = ""
            SERVICES                   = ""
            CARRIER_PARCEL_TYPE        = ""
            CARRIER_TYPE               = "DHL Standardpaket"
            IDCODE                     = $idCode
            EST_SHIP_DATE              = ""
            MESSAGE_COMMENT            = ""
            CUSTOMER_COMMENT           = ""
            INVOICE_NUMBER             = ""
            ORDER_DATE                 = $orderDate
            STOCK_LOCATION_ID          = $stockId
            STOCK_LOCATION_NAME        = $stockName
            SOURCE_STOCK_LOCATION_ID   = $sourceId
            SOURCE_STOCK_LOCATION_NAME = $sourceName
        }
        return $outputString
    }


    # if there is at least one item, add selected locations to the output
    if (($items.Count -ge 1) -and ($selectedStockLocationName1) -and ($selectedSourceLocationName1) -and ($selectedPackageType1)) {
        $output += GetCSVOutputString $items[0] $stockLocationId1 $selectedStockLocationName1 $sourceLocationId1 $selectedSourceLocationName1 $selectedPackageType1
    }

    # work on the optional parameters if order has more than one item and locations are selected on TeamCity
    if (($items.Count -ge 2) -and ($selectedStockLocationName2) -and ($selectedSourceLocationName2) -and ($selectedPackageType2)) {
        Write-Host "Type of the 2nd package: '$selectedPackageType2'"
        $stockLocationId2 = ($stockJson | Where-Object { $_.name -eq $selectedStockLocationName2 }).id
        Write-Host "Found stock location ID for the 2nd item: '$stockLocationId2'"
        $sourceLocationId2 = ($sourceJson | Where-Object { $_.name -eq $selectedSourceLocationName2 }).id
        Write-Host "Found source location ID for the 2nd item: '$sourceLocationId2'"

        $output += GetCSVOutputString $order.ITEMS.ITEM[1] $stockLocationId2 $selectedStockLocationName2 $sourceLocationId2 $selectedSourceLocationName2 $selectedPackageType2

        if (($items.Count -ge 3) -and ($selectedStockLocationName3) -and ($selectedSourceLocationName3) -and ($selectedPackageType3)) {
            Write-Host "Type of the 3rd package: '$selectedPackageType3'"
            $stockLocationId3 = ($stockJson | Where-Object { $_.name -eq $selectedStockLocationName3 }).id
            Write-Host "Found stock location ID for the 3rd item: '$stockLocationId3'"
            $sourceLocationId3 = ($sourceJson | Where-Object { $_.name -eq $selectedSourceLocationName3 }).id
            Write-Host "Found source location ID for the 3rd item: '$sourceLocationId3'"

            $output += GetCSVOutputString $order.ITEMS.ITEM[2] $stockLocationId3 $selectedStockLocationName3 $sourceLocationId3 $selectedSourceLocationName3 $selectedPackageType3
        }
    }
            
    $output | Export-Csv -Path $outputFileFullPath -NoTypeInformation -Delimiter ";"

    Write-Host "Salesreport file $outputFileFullPath generated successfully" 
}

catch [Exception] {
    
    $errorMessage = "FAILED: $_"
    Write-Error $errorMessage
}
