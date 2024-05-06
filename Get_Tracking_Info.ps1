$bearer_tok = Get-Content ./auth/bearer.txt

$headers = @{
    "transId"= "string"
    "transactionSrc"= "1000"
    "Authorization"= "Bearer " + $bearer_tok
}

$query = @{
    locale = "en_US"
    returnSignature= "false"
    returnMilestones= "false"
  }

$shipment = "1Z3124Y81515962179"

$url = "https://onlinetools.ups.com/api/track/v1/details/"+$shipment

$result = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -Body $query

$last_scan = $result.trackResponse.shipment.package

Write-Output $last_scan

#Label Created
<# $label_created = $result.trackResponse.shipment.package.activity
$label_created = $result.trackResponse.shipment.package.activity[$label_created.Length-1].date
$label_created = [datetime]::parseexact($label_created, 'yyyyMMdd', $null).ToString('MM/dd/yyyy') #>

#Current Status
#$curr_status = $result.trackResponse.shipment.package.currentStatus.description

#Shipped Date
<# $label_created = $result.trackResponse.shipment.package.activity
$label_created = $result.trackResponse.shipment.package.activity[$label_created.Length-2].date
$label_created = [datetime]::parseexact($label_created, 'yyyyMMdd', $null).ToString('MM/dd/yyyy') #>

#Last Scan Date/Time


#Last Scan Location




#Delivery Date
#$deliv_date = $result.trackResponse.shipment.package.deliveryDate.date
#$deliv_date = [datetime]::parseexact($deliv_date, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')

#Ship to
<# $to_address1 = $result.trackResponse.shipment.package.packageAddress[1].address.addressLine1
$to_address1 = $to_address1.Substring(0,$to_address1.Length-1)
$to_address2 = $result.trackResponse.shipment.package.packageAddress[1].address.addressLine2
$city = $result.trackResponse.shipment.package.packageAddress[1].address.city
$state = $result.trackResponse.shipment.package.packageAddress[1].address.stateProvince
$zip = $result.trackResponse.shipment.package.packageAddress[1].address.postalCode
$ship_to = ""
if($to_address2.Length -eq 0){
    $ship_to = "$to_address1 $city, $state $zip"
}else{
    $ship_to = "$to_address1 $to_address2 $city, $state $zip"
} #>

#Reference
#$reference = $result.trackResponse.shipment.package.referenceNumber[0].number

#Ship Service
#ship_service = $result.trackResponse.shipment.package.service.description