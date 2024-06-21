$account_number = Get-Content .\auth\account_info.txt
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



$desktop = @(
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

function create_packages($pc, $name, $dept) {
    
    if ($pc -eq "Laptop") {
        $packages = $laptop
        foreach ($item in $packages) {
            $item.ReferenceNumber[0].Value = [string]::Format($item.ReferenceNumber[0].Value, $name)
            $item.ReferenceNumber[1].Value = [string]::Format($item.ReferenceNumber[1].Value, $dept)
        }
        
        Return $packages

    }
    elseif ($pc -eq "Desktop") {
        $packages = $desktop
        foreach ($item in $packages) {
            $item.ReferenceNumber[0].Value = [string]::Format($item.ReferenceNumber[0].Value, $name)
            $item.ReferenceNumber[1].Value = [string]::Format($item.ReferenceNumber[1].Value, $dept)
        }

        Return $packages

    }
    else {
        Write-Error "ERROR IN CREATE PACKAGES. DID NOT SELECT LAPTOP OR DESKTOP"
    }
}


function ship_from($miles_or_roman, $name, $number, $address1, $address2, $city, $state, $zip) {
    if ($miles_or_roman -eq "Miles") {
        Return $from_miles
    }
    elseif ($miles_or_roman -eq "Roman") {
        Return $from_roman
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