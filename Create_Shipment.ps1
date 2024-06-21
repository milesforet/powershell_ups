$bearer_tok = Get-Content .\auth\bearer.txt

$ship_service = @{
    "Next Day Air" = "01"
    "2nd Day Air" = "02"
    "Ground" = "03"
    "3 Day Select" = "12"
    "Next Day Air Saver" = "13"
    "UPS Next Day Air Early" = "14"
}
 

$headers = @{
    "Authorization" = "Bearer $bearer_tok"
    "Content-Type" = "application/json"
    "transId"= "string"
    "transactionSrc"= "1000"
}

$url = "https://wwwcie.ups.com/api/shipments/v2403/ship?additionaladdressvalidation=string"
#$url = "https://onlinetools.ups.com/api/shipments/v2403/ship?additionaladdressvalidation=string"

. ./package_info.ps1

$query = @"
    <View>
        <Query>
            <Where>
                <Eq><FieldRef Name='Status'/><Value Type='Text'>Create Label</Value></Eq>
            </Where>
        </Query>
    </View>
"@

try {
    Connect-PnPOnline -url "https://abs79.sharepoint.com/sites/testing-miles" -interactive
    $sp_list = Get-PnPListItem -List "Create Labels" -Query $query
}
catch {
    Write-Output "Failed to connect to sharepoint or get sharepoint data"
    exit
}

if($sp_list.length -eq 0){
    Write-Output "No shipments to create"
    exit
}

foreach($item in $sp_list){

    try {
        $shipper = ship_from -miles_or_roman $item["ShipFrom"]

    $ship_to_info = @{
        name = $item["Person"]."LookupValue"
        number = $item["PhoneNumber"]
        address1 = $item["AddressLine1"]
        address2 = $item["Address2"]
        city = $item["City"]
        state = $item["State_x0028_Abbr_x0029_"]
        zip = $item["ZipCode"]
        email = $item["Person"].Email
    }

    $ship_to = ship_to @ship_to_info

    $packages = create_packages -pc $item["Laptop_x002f_Desktop"] -name $item["Person"]."LookupValue" -dept $item["Department"]

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
                                AccountNumber = $account_number
                            }
                        }
                    )
                }
                Service = @{
                    Code = $ship_service[$item["ShipService"]]
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

     
    $response = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body $json_body
     
    $response = $response.content | ConvertFrom-Json
    $shipping_result = $response.ShipmentResponse.ShipmentResults.PackageResults
    $count = 1

    $tracking_nums = @()
    foreach($package in $shipping_result){
        $tracking_nums += $package.TrackingNumber
        $label = $package.ShippingLabel.GraphicImage
        [IO.File]::WriteAllBytes("images\$count.jpg", [Convert]::FromBase64String($label))
        $count++
    }

    
    $date = Get-Date -UFormat %m-%d-%y
    $name = $item["Person"]."LookupValue" -replace '\s',''
    python .\convert_pdf.py "$date-$name"
    Remove-item .\images\*.jpg 

    Add-PnPFile -Path "images\$date-$name.pdf" -Folder "Labels"

    $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"Status" = "Label Created"}

    $tracking_nums = ($tracking_nums -join "`n") 
    $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"TrackingNumbers" = $tracking_nums}

    }
    catch {
        Write-Error $_
        $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"Status" = "Failed"; "Notes"= $_}
    }
    
}