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

Function Call-Api($tracking_num){

    $url = "https://onlinetools.ups.com/api/track/v1/details/"+$tracking_num

    $data = @{
        label_created=""
        curr_status="N/A"
        shipped_date="N/A"
        last_scan="N/A"
        last_scan_local="N/A"
        deliv_date="N/A"
        ship_to=""
        reference=""
        ship_service=""
    }

    $result = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -Body $query

    #Label Created
    $data.label_created = $result.trackResponse.shipment.package.activity
    $data.label_created = $result.trackResponse.shipment.package.activity[$label_created.Length-1].date
    $data.label_created = [datetime]::parseexact($data.label_created, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')

    #Current Status
    $data.curr_status = $result.trackResponse.shipment.package.currentStatus.description

    if($data.curr_status -ne "Shipment Ready for UPS"){

        #Shipped Date
        $data.shipped_date = $result.trackResponse.shipment.package.activity
        $data.shipped_date = $result.trackResponse.shipment.package.activity[$shipped_date.Length-2].date
        $data.shipped_date = [datetime]::parseexact($data.shipped_date, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')

        #Last Scan Date/Time
        $last_scan = $result.trackResponse.shipment.package.activity[0].date
        $last_scan = [datetime]::parseexact($last_scan, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')
        $last_time = $result.trackResponse.shipment.package.activity[0].time
        $last_time = [datetime]::parseexact($last_time, 'HHmmss', $null).ToString('hh:mm tt')
        $data.last_scan = "$last_scan $last_time"

        #Last Scan Location
        $data.last_scan_local = $result.trackResponse.shipment.package.activity[0].location.address.city
        $data.last_scan_local = $data.last_scan_local+ ", "  + $result.trackResponse.shipment.package.activity[0].location.address.stateProvince
    }

    if($data.curr_status -eq "Delivered"){
        #Delivery Date
        $data.deliv_date = $result.trackResponse.shipment.package.deliveryDate.date
        $data.deliv_date = [datetime]::parseexact($data.deliv_date, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')
    }

    #Ship to
    $to_address1 = $result.trackResponse.shipment.package.packageAddress[1].address.addressLine1
    $to_address1 = $to_address1.Substring(0,$to_address1.Length-1)
    $to_address2 = $result.trackResponse.shipment.package.packageAddress[1].address.addressLine2
    $city = $result.trackResponse.shipment.package.packageAddress[1].address.city
    $state = $result.trackResponse.shipment.package.packageAddress[1].address.stateProvince
    $zip = $result.trackResponse.shipment.package.packageAddress[1].address.postalCode
    $data.ship_to = ""

    if($to_address2.Length -eq 0){
        $data.ship_to = "$to_address1 $city, $state $zip"
    }else{
        $data.ship_to = "$to_address1 $to_address2 $city, $state $zip"
    }

    #Reference
    $data.reference = $result.trackResponse.shipment.package.referenceNumber[0].number

    #Ship Service
    $data.ship_service = $result.trackResponse.shipment.package.service.description

    Return $data
}


#CONNECT TO SHAREPOINT
Connect-PnPOnline -url "https://70rspw.sharepoint.com/sites/Testing" -interactive

#GET SHAREPOINT LIST
$info = get-pnplistitem -list tracking

for($i=1; $i -le $info.Length; $i++){
    $id = $info[$i-1].Id
    $new_data = Call-Api($info[$i-1]["Title"])

    Set-PnPListItem -List "Tracking" -Identity $id -Values @{"LabelCreated" = $new_data.label_created; "Status"=$new_data.curr_status; "ShippedDate"=$new_data.shipped_date;
    "LastScanDate_x002f_Time"=$new_data.last_scan; "LastScanLocation"=$new_data.last_scan_local;  "ShippingTo"=$new_data.ship_to; "DeliveryDate"=$new_data.deliv_date; 
    "Reference"=$new_data.reference; "ShipService"=$new_data.ship_service}
}