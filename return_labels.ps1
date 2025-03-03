$sp_site_url = "https://abs79.sharepoint.com/sites/ITLogsandAudits"
$sp_list_name = "Equipment Returns"
$creds = Get-AbsCred -credName "service-equip@abskids.net"

$ship_service = @{
    "Next Day Air" = "01"
    "2nd Day Air" = "02"
    "Ground" = "03"
    "3 Day Select" = "12"
    "Next Day Air Saver" = "13"
    "Next Day Air Early" = "14"
}

function create_return($boxes, $name, $returnType){

    
    [System.Collections.ArrayList]$packages = @()

    foreach($item in $boxes){

        $item = $item -replace "x", ""
        Write-Host $item

        $length = $item.SubString(0,2)
        $width = $item.SubString(2,2)
        $height = $item.SubString(4,1)
        Write-Host


        $boxes = @{
                Packaging = @{
                    Code = "02"                        
                }
                Dimensions = @{
                    UnitOfMeasurement = @{
                        Code = "IN"
                    }
                    Length = $length
                    Width = $width
                    Height = $height
                }
                PackageWeight = @{
                    UnitOfMeasurement =  @{
                        Code = "LBS"
                    }
                    Weight = ""
                }
                ReferenceNumber = @(
                    @{
                        Value = ""
                    },
                    @{
                        Value = "Bill to IT"
                    }
                )
            }


        if($returnType -eq "From" -and $length -eq 28){
            $boxes.ReferenceNumber[0].Value = "$name Equip Return"
            $boxes.PackageWeight["Weight"] = "12"
        }elseif ($returnType -eq "From") {
            $boxes.ReferenceNumber[0].Value = "$name Equip Return"
            $boxes.PackageWeight["Weight"] = "8"
        }
        elseif($returnType -eq "To"){
            $boxes.ReferenceNumber[0].Value = "$name Return Boxes"
            $boxes.PackageWeight["Weight"] = "1"
        }


        [void]$packages.Add($boxes)
    }

    return $packages

}


try{
    Connect-PnPOnline -url $sp_site_url -Credentials $creds -ClientId "e41d925a-fe12-4c7f-9675-87a1e5a04e7d"
}catch{
    Write-Logs "Failed to connect to Sharepoint site"
    exit
}

$query = @"
<View>
    <Query>
        <Where>
            <Eq><FieldRef Name='Status'/><Value Type='Create Labels'>Create Labels</Value></Eq>
        </Where>
    </Query>
</View>
"@

$sp_list

try{
    $sp_list = Get-PnPListItem -List $sp_list_name -Query $query
}catch{
    Write-Logs "failed to get data from sharepoint list"
    exit
}

if($sp_list.Length -eq 0){
    Write-Host "No return labels to create"
    exit
}

UPS_Bearer

$bearer_tok = Get-Content C:\Users\MilesForet\Documents\Automations\UPS_Powershell\auth\bearer.txt

:outer foreach($item in $sp_list){

    #checks for assigned
    if($null -eq $item["Assigned"] -or $item["Assigned"] -eq ""){
        Set-PnPListItem -List $sp_list_name -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= "Assigned field is empty"}
        continue
    }

    #saves info in hash
    $ship_to_info = @{
        name = $item["Title"]
        email = $item["PersonalEmail"]
        number = $item["PhoneNumber"]
        address1 = $item["AddressLine1"]
        address2 = $item["AddLine2"]
        city = $item["City"]
        state = $item["State_x0028_Abbr_x0029_"]
        zip = $item["ZipCode"]
        ShipService = $item["ShipService"]
        boxes = $item["Boxes"]
        ShipFrom = $item["Assigned"]
        LabelsNeeded = $item["LabelsNeeded"]
    }


    #data validation
    $ship_to_info.keys | ForEach-Object{

        if(($ship_to_info[$_] -eq $null -or $ship_to_info[$_] -eq "") -and $_ -ne "address2"){
            $i = Set-PnPListItem -List $sp_list_name -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= "$_ field is missing"}
            continue outer
        }
    }

    [System.Collections.ArrayList]$tracking_nums = @()
    $count = 1

    [System.Collections.ArrayList]$addressArray = @()
        [void]$addressArray.Add($ship_to_info.address1)
    if($ship_to_info.address2){
        [void]$addressArray.Add($ship_to_info.address2)
    }

    $shipper =  @{
        Name = "ABS Kids"
        AttentionName = $ship_to_info.name
        CompanyDisplayableName = "ABS Kids"
        EMailAddress = $ship_to_info.email
        Phone = @{
            Number = $ship_to_info.number
        }
        ShipperNumber = "3124Y8"
        Address = @{
            AddressLine = $addressArray
            City = $ship_to_info.city
            StateProvinceCode = $ship_to_info.state
            PostalCode = $ship_to_info.zip
            CountryCode = "US"
        }
    }


    $ship_to = ship_from -miles_or_roman $item["Assigned"]

    $packages = create_return -boxes $ship_to_info.boxes -name $ship_to_info.name -returnType "From"

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


    
    #$url = "https://wwwcie.ups.com/api/shipments/v2403/ship?additionaladdressvalidation=string"
    $url = "https://onlinetools.ups.com/api/shipments/v2403/ship?additionaladdressvalidation=string"

    

    $response = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body $json_body -SkipHttpErrorCheck

    if($response.StatusCode -ne 200){
        $i = Set-PnPListItem -List $sp_list_name -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= "Error creating return labels- $($response.Content)."}
        exit
    }

    $response = $response.content | ConvertFrom-Json

    $shipping_result = $response.ShipmentResponse.ShipmentResults.PackageResults

    foreach($package in $shipping_result){
        $tracking_nums += $package.TrackingNumber
        $label = $package.ShippingLabel.GraphicImage
        [IO.File]::WriteAllBytes("C:\Users\MilesForet\Documents\Automations\UPS_Powershell\images\$count.jpg", [Convert]::FromBase64String($label))
        $count++
    }
    
    if($item["LabelsNeeded"] -eq "To and From User"){

        $shipper = ship_from -miles_or_roman $ship_to_info.ShipFrom
        $ship_to = ship_to @ship_to_info
        $packages = create_return -boxes $ship_to_info.boxes -name $ship_to_info.name -returnType "To"

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

        $response = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body $json_body -SkipHttpErrorCheck

        if($response.StatusCode -ne 200){
            $i = Set-PnPListItem -List $sp_list_name -Identity $item.Id -Values @{"Status" = "Labels Failed"; "LabelCreationNotes"= "Error creating return labels- $($response.Content)."}
            exit
        }
        

        $response = $response.content | ConvertFrom-Json

        $shipping_result = $response.ShipmentResponse.ShipmentResults.PackageResults


        foreach($package in $shipping_result){
            $tracking_nums += $package.TrackingNumber
            $label = $package.ShippingLabel.GraphicImage
            [IO.File]::WriteAllBytes("C:\Users\MilesForet\Documents\Automations\UPS_Powershell\images\$count.jpg", [Convert]::FromBase64String($label))
            $count++
        }
    }

    $date = Get-Date -UFormat %m-%d-%y
    python .\convert_pdf.py "$date Return Boxes - $($ship_to_info.name)"

    Remove-Item ./images/*.jpg

    $i = Add-PnPFile -Path "C:\Users\MilesForet\Documents\Automations\UPS_Powershell\images\$date Return Boxes - $($ship_to_info.name).pdf" -Folder "Labels"
    $trackingNumbersString = ($tracking_nums -join "`n") 
    
    $i = Set-PnPListItem -List $sp_list_name -Identity $item.Id -Values @{"Status" = "Labels Created"; "TrackingNumbers" = $trackingNumbersString;}

    $i = Set-PnPListItem -list $sp_list_name -Identity $item.Id -Values @{"File" = "$sp_site_url/Labels/$date Return Boxes - $($ship_to_info.name).pdf, Return Boxes - $($ship_to_info.name)"}

}