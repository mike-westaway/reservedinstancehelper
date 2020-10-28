# Install-Module -Name Az.Reservations -Scope CurrentUser
# Install-Module ImportExcel -scope CurrentUser

param([string]$Filename = 'ReservedInstanceRequests.xlsx')

function TestFileLock ($FilePath ){
    $FileLocked = $false
    $FileInfo = New-Object System.IO.FileInfo $FilePath
    trap {Set-Variable -name Filelocked -value $true -scope 1; continue}
    $FileStream = $FileInfo.Open( [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None )
    if ($FileStream) {$FileStream.Close()}
    $FileLocked
}

function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref) | out-null
    # Sometime Excel won't shut down immediately - this short sleep fixed the problem 
    Start-Sleep 1
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers() 
}

$ExcelFile = Join-Path -Path $pwd -ChildPath $filename

if (TestFileLock $ExcelFile) {
    write-host "Excel file $ExcelFile is open, close and try again"
    return
}

write-host "Opening excel file, $ExcelFile ..."

$objExcel = New-Object -ComObject Excel.Application  
$WorkBook = $objExcel.Workbooks.Open($ExcelFile)  

# There is an expected format to the workbook, the rows should be on a sheet called Orders 
$WorkSheet = $WorkBook.Sheets.Item("Orders")  
$totalNoOfRecords = ($WorkSheet.UsedRange.Rows).count  

$WorkBookUpdated = $false

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
      Purchased = $WorkSheet.Range("M$rowN").Text
    }

    #debug
    #Write-Host $row

    ###
    # Calculate the Reservation Order
    ###

    if ($row.Purchased -eq "Y") {
        Write-Host "Row $rowN has already been Purchased, skipping"
    }
    else {
        if ($($row.Scope) -eq "Single") {
            $scriptLine = "az reservations reservation-order calculate --sku $($row.Sku) --location $($row.AzureRegion) --reserved-resource-type $($row.ResourceType) --billing-scope $($row.BillingSubscriptionId) --term $($row.Term) --billing-plan $($row.Plan) --quantity $($row.Quantity) --applied-scope-type $($row.Scope) --applied-scope $($row.SingleScopeSubscriptionId) --display-name $($row.ReservationName)"
        }
        else {
            $scriptLine = "az reservations reservation-order calculate --sku $($row.Sku) --location $($row.AzureRegion) --reserved-resource-type $($row.ResourceType) --billing-scope $($row.BillingSubscriptionId) --term $($row.Term) --billing-plan $($row.Plan) --quantity $($row.Quantity) --applied-scope-type $($row.Scope) --display-name $($row.ReservationName)"
        }

        Write-Host $scriptLine

        $ret = Invoke-Expression $scriptLine

        if (!$ret) {
            Write-Host "Error calculate script failed"
        }
        else {

            # script succeeded, convert the result object into JSON
            # result object is an array of strings that are JSON when concatenated back together..

            $retJson = ConvertFrom-Json([string]::Concat($ret))
            
            Write-Host "Order Id $($retJson.properties.reservationOrderId)"

            ###
            # Purchase the reservation order
            ###

            if ($($row.Scope) -eq "Single") {
                #$scriptLine2 = "az reservations reservation-order purchase --reservation-order-id $($retJson.properties.reservationOrderId) --sku $($row.Sku) --location $($row.AzureRegion) --reserved-resource-type $($row.ResourceType) --billing-scope $($row.BillingSubscriptionId) --term $($row.Term) --billing-plan $($row.Plan) --quantity $($row.Quantity) --applied-scope-type $($row.Scope) --applied-scope $($row.SingleScopeSubscriptionId) --display-name $($row.ReservationName)"
                $scriptLine2 = "New-AzReservation -ReservationOrderId $($retJson.properties.reservationOrderId) -Sku $($row.Sku) -Location $($row.AzureRegion) -ReservedResourceType $($row.ResourceType) -BillingScopeId $($row.BillingSubscriptionId) -Term $($row.Term) -BillingPlan  $($row.Plan) -Quantity $($row.Quantity) -AppliedScopeType $($row.Scope) -AppliedScope $($row.SingleScopeSubscriptionId) -DisplayName $($row.ReservationName)"
            }
            else {
                #$scriptLine2 = "az reservations reservation-order purchase --reservation-order-id $($retJson.properties.reservationOrderId) --sku $($row.Sku) --location $($row.AzureRegion) --reserved-resource-type $($row.ResourceType) --billing-scope $($row.BillingSubscriptionId) --term $($row.Term) --billing-plan $($row.Plan) --quantity $($row.Quantity) --applied-scope-type $($row.Scope) --display-name $($row.ReservationName)"
                $scriptLine2 = "New-AzReservation -ReservationOrderId  $($retJson.properties.reservationOrderId) -Sku $($row.Sku) -Location $($row.AzureRegion) -ReservedResourceType $($row.ResourceType) -BillingScopeId $($row.BillingSubscriptionId) -Term $($row.Term) -BillingPlan  $($row.Plan) -Quantity $($row.Quantity) -AppliedScopeType $($row.Scope) -DisplayName $($row.ReservationName)"
            }

            Write-Host $scriptLine2

            $ret = Invoke-Expression $scriptLine2

            if (!$ret) {
                Write-Host "Error purchase script failed"
            }
            else {
                # az returns JSON which Invoke-Expression splits over multiple lines
                #$retJson = ConvertTo-Json(ConvertFrom-Json([string]::Concat($ret)))
                # New-AzReservation returns a PSObject
                $retJson = ConvertTo-Json($ret)

                Write-Host "Purchase script returned $retJson)"
            }

            # Mark the row as Purchased - currently regardless of the outcome from the 2nd Invoke-Expression as this can throw errors due to Role Assigments being async
            $WorkSheet.Range("M$rowN").Value = "Y"
            Write-Host "Row $rowN marked as Purchased" 
            $WorkBookUpdated = $true
        }
    }
}

# Free up all the COM Objects so that Excel can close down quickly and cleanly

if ($WorkBookUpdated) {
    write-host "Saving $ExcelFile"
    $WorkBook.Save()
}
$WorkBook.Close($false)
$objExcel.Quit()

Release-Ref($WorkSheet)
Release-Ref($WorkBook)
Release-Ref($objExcel)

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

