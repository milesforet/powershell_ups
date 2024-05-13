#API CALL FUNCTION THAT RETURNS A PS CUSTOM OBJECT WITH ALL TRACKING INFO
function Get-Shipping-Info($tracking_num){

    $url = "https://onlinetools.ups.com/api/track/v1/details/$tracking_num"

    $tracking_info = @{
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

    $api_response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -Body $query

    #Label Created
    $tracking_info.label_created = $api_response.trackResponse.shipment.package.activity
    $tracking_info.label_created = $api_response.trackResponse.shipment.package.activity[$label_created.Length-1].date
    $tracking_info.label_created = [datetime]::parseexact($tracking_info.label_created, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')

    #Current Status
    $tracking_info.curr_status = $api_response.trackResponse.shipment.package.currentStatus.description

    if($tracking_info.curr_status -ne "Shipment Ready for UPS" -and $tracking_info.curr_status -ne "Shipment Canceled" ){

        #Shipped Date
        $tracking_info.shipped_date = $api_response.trackResponse.shipment.package.activity
        $tracking_info.shipped_date = $api_response.trackResponse.shipment.package.activity[$shipped_date.Length-2].date
        $tracking_info.shipped_date = [datetime]::parseexact($tracking_info.shipped_date, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')

        #Last Scan Date/Time
        $last_scan = $api_response.trackResponse.shipment.package.activity[0].date
        $last_scan = [datetime]::parseexact($last_scan, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')
        $last_time = $api_response.trackResponse.shipment.package.activity[0].time
        $last_time = [datetime]::parseexact($last_time, 'HHmmss', $null).ToString('hh:mm tt')
        $tracking_info.last_scan = "$last_scan $last_time"

        #Last Scan Location
        $tracking_info.last_scan_local = $api_response.trackResponse.shipment.package.activity[0].location.address.city
        $tracking_info.last_scan_local = $tracking_info.last_scan_local+ ", "  + $api_response.trackResponse.shipment.package.activity[0].location.address.stateProvince
    }

    if($tracking_info.curr_status -eq "Delivered"){
        #Delivery Date
        $tracking_info.deliv_date = $api_response.trackResponse.shipment.package.deliveryDate.date
        $tracking_info.deliv_date = [datetime]::parseexact($tracking_info.deliv_date, 'yyyyMMdd', $null).ToString('MM/dd/yyyy')
    }

    #Ship to
    $to_address1 = $api_response.trackResponse.shipment.package.packageAddress[1].address.addressLine1
    $to_address1 = $to_address1.Substring(0,$to_address1.Length-1)
    $to_address2 = $api_response.trackResponse.shipment.package.packageAddress[1].address.addressLine2
    $city = $api_response.trackResponse.shipment.package.packageAddress[1].address.city
    $state = $api_response.trackResponse.shipment.package.packageAddress[1].address.stateProvince
    $zip = $api_response.trackResponse.shipment.package.packageAddress[1].address.postalCode
    $tracking_info.ship_to = ""

    if($to_address2.Length -eq 0){
        $tracking_info.ship_to = "$to_address1 $city, $state $zip"
    }else{
        $tracking_info.ship_to = "$to_address1 $to_address2 $city, $state $zip"
    }

    #Reference
    $tracking_info.reference = $api_response.trackResponse.shipment.package.referenceNumber.number[2]

    #Ship Service
    $tracking_info.ship_service = $api_response.trackResponse.shipment.package.service.description

    Return $tracking_info
}
#FUNCTION TO VOID SHIPMENT. TAKES TRACKING NUMBER AS PARAMETER
function Void-Shipment($tracking_num){

    $url = "https://onlinetools.ups.com/api/shipments/v2403/void/cancel/"+$tracking_num+"?trackingnumber="+$tracking_num

    $bearer_tok = Get-Content ./auth/bearer.txt

    $headers = @{
    "Authorization"= "Bearer $bearer_tok"
    "transId"= "string"
    "transactionSrc"= "1000"
    }

    #MAKE REQUEST TO VOID SHIPMENT
    Invoke-RestMethod -Uri $url -Method Delete -Headers $headers
}

./Bearer_Token.ps1

<#
    -Connecting to Sharepoint, getting all items in list that have "Void Shipment" selected
    -Iterate through list if it's not empty. Call "Void-Shipment" function on all shipments, which will void the shipments
    -Updates list item to "Shipment Voided" so it doesn't try to void them again
#>
Connect-PnPOnline -url "https://70rspw.sharepoint.com/sites/Testing" -interactive

try {
    $void_shipment = @"
    <View>
        <Query>
            <Where>
                <Eq><FieldRef Name='Void'/><Value Type='Text'>Void Shipment</Value></Eq>
            </Where>
        </Query>
    </View>
"@
    $sp_list =Get-PnPListItem -list tracking -Query $void_shipment

    if($sp_list){

        foreach($item in $sp_list){

            #Void-Shipment($item["Title"])
            Write-Host $item["Title"] "- Voided!"
            $i = Set-PnPListItem -List "Tracking" -Identity $item.Id -Values @{"Void"="Shipment Voided"}
        }
    }else{
            Write-Output "No shipments to void"
        }

}
catch {
    Write-host -f red "Encountered Error:"$_.Exception.Message
}



<#
    -Get all items in list where the "Status" is not "Deliverd" or "Shipment Cancelled"
    -Iterate through list. Check format to ensure it's correct length and starts with "1Z"
    -Call UPS Tracking endpoint for shipping info on each package. Update the list with updated tracking information
#>
try {
    $active_package_query = @"
    <View>
        <Query>
            <Where>
                <And>
                    <Neq><FieldRef Name='Status'/><Value Type='Text'>Delivered</Value></Neq>
                    <Neq><FieldRef Name='Status'/><Value Type='Text'>Shipment Canceled</Value></Neq>
                </And>
            </Where>
        </Query>
    </View>
"@
    $sp_list = Get-PnPListItem -list tracking -Query $active_package_query

    foreach($item in $sp_list){

        if($item["Title"] -clike "1Z*" -and $item["Title"].Length -eq 18){

            $package_info = Get-Shipping-Info($item["Title"])

            $i = Set-PnPListItem -List "Tracking" -Identity $item.Id -Values @{"LabelCreated" = $package_info.label_created; "Status"=$package_info.curr_status; "ShippedDate"=$package_info.shipped_date;
            "LastScanDate_x002f_Time"=$package_info.last_scan; "LastScanLocation"=$package_info.last_scan_local;  "ShippingTo"=$package_info.ship_to; "DeliveryDate"=$package_info.deliv_date; 
            "Reference"=$package_info.reference; "ShipService"=$package_info.ship_service}

            Write-Host $item["Title"] "has been updated"
        
        }else{
            Write-Host "----Id:"$item.ID $item["Title"] "format incorrect. Not a UPS tracking number----"
            $i = Set-PnPListItem -List "Tracking" -Identity $item.Id -Values @{"Status"= "Not Valid UPS Shipment"}
        }
  }
}catch{
    Write-host -f red "Encountered Error:"$_.Exception.Message
}