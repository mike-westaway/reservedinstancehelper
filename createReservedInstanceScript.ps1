# Install-Module ImportExcel -scope CurrentUser

param([string]$Filename = 'ReservedInstanceRequests.xlsx')

$ExcelFile = Join-Path -Path $pwd -ChildPath $filename

write-host "Opening excel file, $ExcelFile ..."

$objExcel = New-Object -ComObject Excel.Application  
$WorkBook = $objExcel.Workbooks.Open($ExcelFile)  

# There is an expected format to the workbook, the rows should be on a sheet called Orders 
$WorkSheet = $WorkBook.Sheets.Item("Orders")  
$totalNoOfRecords = ($WorkSheet.UsedRange.Rows).count  

for ( $rowN = 2; $rowN -le $totalNoOfRecords; $rowN++ ) {

    $row = [pscustomobject][ordered]@{
      ReservationName = $WorkSheet.Range("A$rowN").Text
      ResourceType = $WorkSheet.Range("B$rowN").Text
      Term = $WorkSheet.Range("C$rowN").Text
      Plan = $WorkSheet.Range("D$rowN").Text
      BillingSubscriptionId = $WorkSheet.Range("E$rowN").Text
      Scope = $WorkSheet.Range("F$rowN").Text
      SingleScopeSubscriptionId = $WorkSheet.Range("G$rowN").Text
      InstanceFlexability = $WorkSheet.Range("H$rowN").Text
      AzureRegion = $WorkSheet.Range("I$rowN").Text
      AutoRenew = $WorkSheet.Range("J$rowN").Text
      Sku = $WorkSheet.Range("K$rowN").Text
      Quantity = $WorkSheet.Range("L$rowN").Text
    }

    #debug
    #Write-Host $row

    ###
    # Calculate the Reservation Order
    ###

    if ($($row.Scope) -eq "Single") {
        $scriptLine = "az reservations reservation-order calculate --sku $($row.Sku) --location $($row.AzureRegion) --reserved-resource-type $($row.ResourceType) --billing-scope $($row.BillingSubscriptionId) --term $($row.Term) --billing-plan $($row.Plan) --quantity $($row.Quantity) --applied-scope-type $($row.Scope) --applied-scope $($row.SingleScopeSubscriptionId) --display-name $($row.ReservationName)"
    }
    else {
        $scriptLine = "az reservations reservation-order calculate --sku $($row.Sku) --location $($row.AzureRegion) --reserved-resource-type $($row.ResourceType) --billing-scope $($row.BillingSubscriptionId) --term $($row.Term) --billing-plan $($row.Plan) --quantity $($row.Quantity) --applied-scope-type $($row.Scope) --display-name $($row.ReservationName)"
    }

    Write-Host $scriptLine

    $ret = Invoke-Expression $scriptLine

    if (!$ret) {
        Write-Host "Error script failed"
    }
    else {

        # script succeeded, convert the result object into JSON
        $retJson = ConvertFrom-Json $ret
        
        Write-Host $retJson.reservationOrderId

        ###
        # Purchase the reservation order
        ###

        if ($($row.Scope) -eq "Single") {
            $scriptLine2 = 'az reservations reservation-order purchase --reservation-order-id $($retJson.reservationOrderId) --sku $($row.Sku) --location $($row.AzureRegion) --reserved-resource-type $($row.ResourceType) --billing-scope $($row.BillingSubscriptionId) --term $($row.Term) --billing-plan $($row.Plan) --quantity $($row.Quantity) --applied-scope-type $($row.Scope) --applied-scope $($row.SingleScopeSubscriptionId) --display-name $($row.ReservationName)'
        }
        else {
            $scriptLine2 = 'az reservations reservation-order purchase --reservation-order-id $($retJson.reservationOrderId) --sku $($row.Sku) --location $($row.AzureRegion) --reserved-resource-type $($row.ResourceType) --billing-scope $($row.BillingSubscriptionId) --term $($row.Term) --billing-plan $($row.Plan) --quantity $($row.Quantity) --applied-scope-type $($row.Scope) --display-name $($row.ReservationName)'
        }

        Write-Host $scriptLine2
    }
}

# Free up all the COM Objects so that Excel can close down quickly and cleanly

$WorkBook.Close($false)
$objExcel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Remove-Variable -Name objExcel

# Notes

# expected return format from az reservations reservation-order calculate..
# https://docs.microsoft.com/en-us/rest/api/reserved-vm-instances/reservationorder/calculate#examples
# {
#   "properties": {
#       "billingCurrencyTotal": {
#         "currencyCode": "USD",
#         "amount": 46
#       },
#       "reservationOrderId": "6d9cec54-7de8-abcd-9de7-80f5d634f2d2",
#       "skuTitle": "Reserved VM Instance, Standard_D1, US West, 1 Year",
#       "skuDescription": "standard_D1",
#       "pricingCurrencyTotal": {
#         "currencyCode": "USD",
#         "amount": 46
#       },
#       "paymentSchedule": [
#         {
#           "dueDate": "2019-05-14",
#           "pricingCurrencyTotal": {
#             "currencyCode": "USD",
#             "amount": 46
#           },
#           "billingCurrencyTotal": {
#             "currencyCode": "EUR",
#             "amount": 40
#           }
#         }
#       ]
#     }
#   }

