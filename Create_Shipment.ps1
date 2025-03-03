$ship_service = @{
    "Next Day Air" = "02"
    "2nd Day Air" = "02"
    "Ground" = "03"
    "3 Day Select" = "12"
    "Next Day Air Saver" = "13"
    "Next Day Air Early" = "14"
}


$creds = Get-AbsCred -credName "service-equip@abskids.net"


#$url = "https://wwwcie.ups.com/api/shipments/v2403/ship?additionaladdressvalidation=string"
$url = "https://onlinetools.ups.com/api/shipments/v2403/ship?additionaladdressvalidation=string"

$query = @"
    <View>
        <Query>
            <Where>
                <Eq><FieldRef Name='Status'/><Value Type='Text'>Create Labels</Value></Eq>
            </Where>
        </Query>
    </View>
"@

try {
    Connect-PnPOnline -url "https://abs79.sharepoint.com/sites/ITLogsandAudits" -Credentials $creds -ClientId "e41d925a-fe12-4c7f-9675-87a1e5a04e7d"
    $sp_list = Get-PnPListItem -List "Equipment Prep" -Query $query
}
catch {
    Write-Logs "Failed to connect to sharepoint or get sharepoint data"
    exit
}

if($sp_list.length -eq 0){
    exit
}


UPS_Bearer

$bearer_tok = Get-Content C:\Users\MilesForet\Documents\Automations\UPS_Powershell\auth\bearer.txt

:outer foreach($item in $sp_list){

    try {
        if($null -eq $item["Assigned"] -or $item["Assigned"] -eq ""){
            $i = Set-PnPListItem -List "Equipment Prep" -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= "Missing Ship From field"}
            continue
        }

        $ship_to_info = @{
            name = $item["Title"]
            number = $item["PhoneNumber"]
            address1 = $item["AddressLine1"]
            address2 = $item["AddressLine2"]
            city = $item["City"]
            state = $item["State_x0028_Abbr_x0029_"]
            zip = $item["ZipCode"]
            ShipService = $item["ShipService"]
            laptop_or_desktop = $item["Laptop_x002f_Desktop"]
            dept = $item["BillTo"]
            ShipFrom = $item["Assigned"]
        }

        if(($ship_to_info.dept).length -gt 24){
            Write-Logs "$($item["Title"]) - Bill to Adjusted due to being too long"
            Write-Logs "Old - $($ship_to_info.dept)"
            $ship_to_info.dept = ($ship_to_info.dept).SubString(0, 24)
            Write-Logs "New - $($ship_to_info.dept)"
            $i = Set-PnPListItem -List "Equipment Prep" -Identity $item.Id -Values @{"LabelCreationNotes"= "$($item["Title"]) - Bill to Adjusted due to being too long"}
        }

        #validate data
        $ship_to_info.keys | ForEach-Object{

            if(($ship_to_info[$_] -eq $null -or $ship_to_info[$_] -eq "") -and $_ -ne "address2"){
                $i = Set-PnPListItem -List "Equipment Prep" -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= "$_ field is missing"}
                continue outer
            }
        }


        $shipper = ship_from -miles_or_roman $ship_to_info.ShipFrom

        $ship_to = ship_to @ship_to_info

        $packages = create_packages -pc $ship_to_info["laptop_or_desktop"] -name $ship_to_info["name"] -dept $ship_to_info["dept"]

        if($packages -eq "Error"){
            Write-Logs "$($item["Title"]) - Error in creating packages"
            $i = Set-PnPListItem -List "Equipment Prep" -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= "Error in creating packages"}
            continue
        }

        $body = @{
            ShipmentRequest = @{
                Request = @{
                    RequestOption = "validate"
                }
                Shipment = @{
                    Shipper = $shipper
            
                    ShipTo = $ship_to
        
                    PaymentInformation = @{
                        ShipmentCharge = @(
                            @{
                                Type = "01"
                                BillShipper = @{
                                    AccountNumber = "3124Y8"
                                }
                            }
                        )
                    }
                    Service = @{
                        Code = $ship_service[$ship_to_info["ShipService"]]
                    }
                    Package = $packages
                }
                LabelSpecification = @{
                    LabelImageFormat = @{
                        Code = "GIF"
                    }
                    LabelStockSize = @{
                        Height = "6"
                        Width = "4"
                    }
                }
            
            }
        }
        
        $json_body = $body | ConvertTo-Json -depth 15

        $headers = @{
            "Authorization" = "Bearer $bearer_tok"
            "Content-Type" = "application/json"
            "transId"= "string"
            "transactionSrc"= "1000"
        }
        
        $response = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body $json_body
        
        $response = $response.content | ConvertFrom-Json
        $shipping_result = $response.ShipmentResponse.ShipmentResults.PackageResults
        $count = 1

        $tracking_nums = @()

        foreach($package in $shipping_result){
            $tracking_nums += $package.TrackingNumber
            $label = $package.ShippingLabel.GraphicImage
            [IO.File]::WriteAllBytes("C:\Users\MilesForet\Documents\Automations\UPS_Powershell\images\$count.jpg", [Convert]::FromBase64String($label))
            $count++
        }

        $count--
        try{
            $date = Get-Date -UFormat %m-%d-%y_%H-%M
            $name = $item["Title"] -replace '\s',''
            python .\convert_pdf.py "$date-$name-$count"
            Remove-item "C:\Users\MilesForet\Documents\Automations\UPS_Powershell\images\*.jpg"

            $i = Add-PnPFile -Path "C:\Users\MilesForet\Documents\Automations\UPS_Powershell\images\$date-$name-$count.pdf" -Folder "Labels"
            $tracking_nums = ($tracking_nums -join "`n") 
            $i = Set-PnPListItem -List "Equipment Prep" -Identity $item.Id -Values @{"Status" = "Labels Created"; "ShippingInfo" = $tracking_nums;}

            try{
                $i = Set-PnPListItem -list "Equipment Prep" -Identity $item.Id -Values @{"File" = "https://abs79.sharepoint.com/sites/ITLogsandAudits/Labels/$date-$name-$count.pdf, $date-$name-$count.pdf"}
            }catch{
                Write-Output "Gives an error, but works. Trying to figure out why Set-PnpListItem doesn't like Hyperlink columns"
            }
            
            Write-Logs "$($item["Title"]) - Labels Created"

            Clear-Variable -Name ship_to_info, body, packages

        }catch{
            Write-Logs "$_ - Error in PDF creation or updating info"
            Clear-Variable -Name ship_to_info, body, packages
            $i = Set-PnPListItem -List "Equipment Prep" -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= $_}
        }

        }
        catch {

            Write-Logs "$_"
            $i = Set-PnPListItem -List "Equipment Prep" -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= $_}
        }
}
