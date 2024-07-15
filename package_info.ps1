$account_number = Get-Content .\auth\account_info.txt

$from_dalton = @{
    Name = "ABS Kids"
    AttentionName = "Dalton Grange"
    CompanyDisplayableName = "ABS Kids"
    EMailAddress = "dgrange@abskids.com"
    Phone = @{
        Number = "3853601314"
    }
    ShipperNumber = $account_number
    Address = @{
        AddressLine = "515 S 700 E", "Ste 2L"
        City = "SALT LAKE CITY"
        StateProvinceCode = "UT"
        PostalCode = "84102"
        CountryCode = "US"
    }
}

$from_roman = @{
    Name = "ABS Kids"
    AttentionName = "Roman Samul"
    CompanyDisplayableName = "ABS Kids"
    EMailAddress = "rsamul@abskids.com"
    Phone = @{
        Number = "8016754623"
    }
    ShipperNumber = $account_number
    Address = @{
        AddressLine = "515 S 700 E", "Ste 2L"
        City = "SALT LAKE CITY"
        StateProvinceCode = "UT"
        PostalCode = "84102"
        CountryCode = "US"
    }
}
 
$from_miles = @{
    Name = "ABS Kids"
    AttentionName = "Miles Foret"
    CompanyDisplayableName = "ABS Kids"
    EMailAddress = "mforet@abskids.com"
    Phone = @{
        Number = "7047272011"
    }
    ShipperNumber = $account_number
    Address = @{
        AddressLine = ,"258 East Garrison Blvd"
        City = "GASTONIA"
        StateProvinceCode = "NC"
        PostalCode = "28054"
        CountryCode = "US"
    }
}


function laptop($name, $dept){
    $laptop = @(
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "24"
            Width = "16"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "14"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Monitors"
            },
            @{
                Value = "Bill to {0}"
            }
        )
    },
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "24"
            Width = "16"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "14"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Monitors"
            },
            @{
                Value = "Bill to {0}"
            }
        )
    },
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "20"
            Width = "10"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "8"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Peripherals"
            },
            @{
                Value = "Bill to {0}"
            }
        )
    },
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "17"
            Width = "13"
            Height = "3"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "6"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Laptop"
            },
            @{
                Value = "Bill to {0}"
            }
        )
        PackageServiceOptions = @{
            DeclaredValue = @{
                CurrencyCode ="USD"
                MonetaryValue = "999"
            }
        }
    }
    )

    foreach ($item in $laptop) {
        $item.ReferenceNumber[0].Value = [string]::Format($item.ReferenceNumber[0].Value, $name)
        $item.ReferenceNumber[1].Value = [string]::Format($item.ReferenceNumber[1].Value, $dept)
    }
    return $laptop
    
}

function desktop($name, $dept){
    $desktop =@(
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "24"
            Width = "16"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "14"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Monitors"
            },
            @{
                Value = "Bill to {0}"
            }
        )
    },
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "24"
            Width = "16"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "14"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Monitors"
            },
            @{
                Value = "Bill to {0}"
            }
        )
    },
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "20"
            Width = "10"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "8"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Peripherals"
            },
            @{
                Value = "Bill to {0}"
            }
        )
    },
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "20"
            Width = "10"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "8"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Desktop"
            },
            @{
                Value = "Bill to {0}"
            }
        )
        PackageServiceOptions = @{
            DeclaredValue = @{
                CurrencyCode = "USD"
                MonetaryValue = "999"
            }
        }
    }
)
    foreach ($item in $desktop) {
        $item.ReferenceNumber[0].Value = [string]::Format($item.ReferenceNumber[0].Value, $name)
        $item.ReferenceNumber[1].Value = [string]::Format($item.ReferenceNumber[1].Value, $dept)
    }
    return $desktop
}

function only_laptop($name, $dept){
    $laptop = @(
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "17"
            Width = "13"
            Height = "3"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "6"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Laptop"
            },
            @{
                Value = "Bill to {0}"
            }
        )
        PackageServiceOptions = @{
            DeclaredValue = @{
                CurrencyCode ="USD"
                MonetaryValue = "999"
            }
        }
    }
    )

    $laptop[0].ReferenceNumber[0].Value = [string]::Format($laptop[0].ReferenceNumber[0].Value, $name)
    $laptop[0].ReferenceNumber[1].Value = [string]::Format($laptop[0].ReferenceNumber[1].Value, $dept)

    return $laptop
    
}

function only_dock($name, $dept){
    $dock = @(
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "10"
            Width = "10"
            Height = "3"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "2"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Dock"
            },
            @{
                Value = "Bill to {0}"
            }
        )
        PackageServiceOptions = @{
            DeclaredValue = @{
                CurrencyCode ="USD"
                MonetaryValue = "999"
            }
        }
    }
    )

    $dock[0].ReferenceNumber[0].Value = [string]::Format($dock[0].ReferenceNumber[0].Value, $name)
    $dock[0].ReferenceNumber[1].Value = [string]::Format($dock[0].ReferenceNumber[1].Value, $dept)
    
    return $dock
}

function only_monitor($name, $dept){
    $monitor = @(
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "24"
            Width = "16"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "20"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Monitor"
            },
            @{
                Value = "Bill to {0}"
            }
        )
    })

    $monitor[0].ReferenceNumber[0].Value = [string]::Format($monitor[0].ReferenceNumber[0].Value, $name)
    $monitor[0].ReferenceNumber[1].Value = [string]::Format($monitor[0].ReferenceNumber[1].Value, $dept)
    
    return $monitor
}

function only_peripheral($name, $dept){
    $peripherals = @(
        @{
            Packaging = @{
                Code = "02"                        
            }
            Dimensions = @{
                UnitOfMeasurement = @{
                    Code = "IN"
                }
                Length = "20"
                Width = "10"
                Height = "6"
            }
            PackageWeight = @{
                UnitOfMeasurement =  @{
                    Code = "LBS"
                }
                Weight = "8"
            }
            ReferenceNumber = @(
                @{
                    Value = "{0} Peripherals"
                },
                @{
                    Value = "Bill to {0}"
                }
            )
        })

    $peripherals.ReferenceNumber[0].Value = [string]::Format($peripherals.ReferenceNumber[0].Value, $name)
    $peripherals.ReferenceNumber[1].Value = [string]::Format($peripherals.ReferenceNumber[1].Value, $dept)
    
    return $peripherals
}

function only_performance_laptop($name, $dept){
    $laptop = @(
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "17"
            Width = "13"
            Height = "5"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "6"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} Laptop"
            },
            @{
                Value = "Bill to {0}"
            }
        )
        PackageServiceOptions = @{
            DeclaredValue = @{
                CurrencyCode ="USD"
                MonetaryValue = "999"
            }
        }
    }
    )

    $laptop[0].ReferenceNumber[0].Value = [string]::Format($laptop[0].ReferenceNumber[0].Value, $name)
    $laptop[0].ReferenceNumber[1].Value = [string]::Format($laptop[0].ReferenceNumber[1].Value, $dept)

    return $laptop
}

function create_packages($pc, $name, $dept) {

    if ($pc -eq "Laptop Package") {
        return laptop -name $name -dept $dept

    }
    elseif ($pc -eq "Desktop Package") {
        return desktop -name $name -dept $dept

    }else{
        $packages = @()

        foreach($item in $pc){

            #TODO: change to switch statement
            if($item -eq "Desktop Package" -or $item -eq "Laptop Package"){
                Return "Error"

            }elseif($item -eq "Dock"){
                $packages += (only_dock -name $name -dept $dept)
            }
            elseif($item -eq "Laptop"){
                $packages += (only_laptop -name $name -dept $dept)
            }
            elseif($item -eq "Monitor 1" -or $item -eq "Monitor 2"){
                $packages += (only_monitor -name $name -dept $dept)
            }
            elseif($item -eq "Peripherals"){
                $packages += (only_peripheral -name $name -dept $dept)
            }
            elseif($item -eq "Performance Laptop"){
                $packages += (only_performance_laptop -name $name -dept $dept)
            }
        }
        return $packages
    }
}


function ship_from($miles_or_roman, $name, $number, $address1, $address2, $city, $state, $zip) {
    if ($miles_or_roman -eq "Miles") {
        Return $from_miles
    }
    elseif ($miles_or_roman -eq "Roman") {
        Return $from_roman
    }
    elseif ($miles_or_roman -eq "Dalton"){
        return $from_dalton
    }
    else {
        #TODO: add customizable shipping address. Will implement later
        Write-Error "Failed in ship_from function. Value has to be Miles or Roman"
    }
}

function ship_to($name, $number, $address1, $address2, $city, $state, $zip, $email) {
    
    $address = @()
    if ($null -ne $address1 -and $null -ne $address2) {
        $address += $address1, $address2
    }
    elseif ($null -ne $address1 -and $null -eq $address2) {
        $address += $address1
    }
    else {
        Write-Output "No addresses"
        exit
    }
    
    
    $ship_to = @{
        Name = "ABS Kids"
        AttentionName = $name
        CompanyDisplayableName = "ABS Kids"
        EMailAddress = $email
        Phone = @{
            Number = $number
        }
        Address = @{
            AddressLine = $address
            City = $city
            StateProvinceCode = $state
            PostalCode = $zip
            CountryCode = "US"
        }
    }

    return $ship_to
}