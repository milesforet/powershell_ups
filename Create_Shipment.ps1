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

:outer foreach($item in $sp_list){

    $item
    try {
        if($item["ShipFrom"] -eq $null -or $item["ShipFrom"] -eq ""){
            $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"Status" = "Failed"; "Notes"= "Missing Ship From field"}
            continue
        }

        $ship_to_info = @{
            name = $item["Title"]
            number = $item["PhoneNumber"]
            address1 = $item["AddressLine1"]
            address2 = $item["Address2"]
            city = $item["City"]
            state = $item["State_x0028_Abbr_x0029_"]
            zip = $item["ZipCode"]
            ShipService = $item["ShipService"]
            laptop_or_desktop = $item["Laptop_x002f_Desktop"]
            dept = $item["Department"]
            ShipFrom = $item["ShipFrom"]
        }

        #validate data
        $ship_to_info.keys | ForEach-Object{

            if(($ship_to_info[$_] -eq $null -or $ship_to_info[$_] -eq "") -and $_ -ne "address2"){
                $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"Status" = "Failed"; "Notes"= "$_ field is missing"}
                continue outer
            }
        }

        $shipper = ship_from -miles_or_roman $ship_to_info["ShipFrom"]

        $ship_to = ship_to @ship_to_info

        $packages = create_packages -pc $ship_to_info["laptop_or_desktop"] -name $ship_to_info["name"] -dept $ship_to_info["dept"]

        if($packages -eq "Error"){
            $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"Status" = "Failed"; "Notes"= "Error in creating packages"}
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
                                    AccountNumber = $account_number
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

        $count--
        try{
            $date = Get-Date -UFormat %m-%d-%y_%H-%M
            $name = $item["Title"] -replace '\s',''
            python .\convert_pdf.py "$date-$name-$count"
            Remove-item .\images\*.jpg 

            $i = Add-PnPFile -Path "images\$date-$name-$count.pdf" -Folder "Labels"
            $tracking_nums = ($tracking_nums -join "`n") 
            $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"Status" = "Label Created"; "TrackingNumbers" = $tracking_nums; "Notes" = "";}

            try{
                $i = Set-PnPListItem -list "Create Labels" -Identity $item.Id -Values @{"File" = "https://abs79.sharepoint.com/sites/testing-miles/Labels/$date-$name-$count.pdf, $date-$name-$count.pdf"}
            }catch{
                Write-Output "Gives an error, but works. Trying to figure out why Set-PnpListItem doesn't like Hyperlink columns"
            }
            Clear-Variable -Name ship_to_info, body, packages

        }catch{
            Write-Error "$_ - Error in PDF creation or updating info"
            Clear-Variable -Name ship_to_info, body, packages
            $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"Status" = "Failed"; "Notes"= $_}
        }

        }
        catch {

            Write-Error "$_ - ${$_.InvocationInfo.ScriptLineNumber}"
            $i = Set-PnPListItem -List "Create Labels" -Identity $item.Id -Values @{"Status" = "Failed"; "Notes"= $_}
        }
}