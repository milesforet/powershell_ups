$from_dalton = @{
    Name = "ABS Kids"
    AttentionName = "Dalton Grange"
    CompanyDisplayableName = "ABS Kids"
    EMailAddress = "dgrange@abskids.com"
    Phone = @{
        Number = "3853601314"
    }
    ShipperNumber = "3124Y8"
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
    ShipperNumber = "3124Y8"
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
    ShipperNumber = "3124Y8"
    Address = @{
        AddressLine = ,"258 East Garrison Blvd"
        City = "GASTONIA"
        StateProvinceCode = "NC"
        PostalCode = "28054"
        CountryCode = "US"
    }
}

$from_warren = @{
    Name = "ABS Kids"
    AttentionName = "Warren Cramer"
    CompanyDisplayableName = "ABS Kids"
    EMailAddress = "wcramer@abskids.com"
    Phone = @{
        Number = "7432576006"
    }
    ShipperNumber = "3124Y8"
    Address = @{
        AddressLine = "257 Gretas Way", "Suite 100"
        City = "KERNERSVILLE"
        StateProvinceCode = "NC"
        PostalCode = "27284"
        CountryCode = "US"
    }
}


New-Variable -Name from_dalton -Value $from_dalton -Scope Script -Force
New-Variable -Name from_miles -Value $from_miles -Scope Script -Force
New-Variable -Name from_roman -Value $from_roman -Scope Script -Force
New-Variable -Name from_warren -Value $from_warren -Scope Script -Force



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
            Weight = "13"
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
            Length = "20"
            Width = "20"
            Height = "5"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "7"
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


function aio_desktop($name, $dept){
    $laptop = @(
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "29"
            Width = "17"
            Height = "6"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "16"
        }
        ReferenceNumber = @(
            @{
                Value = "{0} AIO Desktop"
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


function ipads($name, $dept){
    $ipads = @(
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
                    Value = "{0} iPads"
                },
                @{
                    Value = "Bill to {0}"
                }
            )
        })

    $ipads.ReferenceNumber[0].Value = [string]::Format($ipads.ReferenceNumber[0].Value, $name)
    $ipads.ReferenceNumber[1].Value = [string]::Format($ipads.ReferenceNumber[1].Value, $dept)
    
    return $ipads
}

function only_desktop($name, $dept){
    $desktop = @(
    @{
        Packaging = @{
            Code = "02"                        
        }
        Dimensions = @{
            UnitOfMeasurement = @{
                Code = "IN"
            }
            Length = "20"
            Width = "20"
            Height = "5"
        }
        PackageWeight = @{
            UnitOfMeasurement =  @{
                Code = "LBS"
            }
            Weight = "7"
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
                CurrencyCode ="USD"
                MonetaryValue = "999"
            }
        }
    }
    )

    $desktop[0].ReferenceNumber[0].Value = [string]::Format($desktop[0].ReferenceNumber[0].Value, $name)
    $desktop[0].ReferenceNumber[1].Value = [string]::Format($desktop[0].ReferenceNumber[1].Value, $dept)

    return $desktop
}


function create_packages($pc, $name, $dept) {

    $packages = @()

    foreach($item in $pc){
        Write-Logs $item

        switch ($item) {
            "Laptop Package" {$packages += (laptop -name $name -dept $dept)}
            "Desktop Package" {$packages += desktop -name $name -dept $dept}
            "Dock" {$packages += (only_dock -name $name -dept $dept)}
            "Laptop" {$packages += (only_laptop -name $name -dept $dept)}
            "Monitor 1" {$packages += (only_monitor -name $name -dept $dept)}
            "Monitor 2" {$packages += (only_monitor -name $name -dept $dept)}
            "Peripherals" {$packages += (only_peripheral -name $name -dept $dept)}
            "Performance Laptop" {$packages += (only_performance_laptop -name $name -dept $dept)}
            "iPads" {$packages += (ipads -name $name -dept $dept)}
            "AIO" {$packages += (aio_desktop -name $name -dept $dept)}
            "Desktop" {$packages += (only_desktop -name $name -dept $dept)}
            Default {return "Error"}
        }
    }
    return $packages
    
}

function ship_from($miles_or_roman, $name, $number, $address1, $address2, $city, $state, $zip) {
    if ($miles_or_roman -eq "Miles") {
        return $from_miles
    }
    elseif ($miles_or_roman -eq "Roman") {
        return $from_roman
    }
    elseif ($miles_or_roman -eq "Dalton"){
        return $from_dalton
    }
    elseif ($miles_or_roman -eq "Warren"){
        return $from_warren
    }
    else {
        #TODO: add customizable shipping address. Will implement later
        Write-Error "Failed in ship_from function."
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
        AttentionName = $name
        Name = "ABS Kids"
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

Function Write-Logs($msg){
    Write-Host $msg
    $date = Get-Date -UFormat %m-%d-%y
    $time = Get-Date -UFormat %I:%M
    $file = (Split-Path $MyInvocation.PSCommandPath -Leaf) -replace ".ps1"
    "[$($time)] $file - $msg" | Out-File -Append -FilePath "C:\Users\MilesForet\Documents\Automations\--Logs--\$date.txt" -Encoding UTF8
} 

function UPS_Bearer{
    $refresh = (Get-Content C:\Users\MilesForet\Documents\Automations\UPS_Powershell\auth\bearer_refresh.txt)
    $refresh = [datetime]::ParseExact($refresh, "MM/dd/yy hh:mm tt", $null)
    if((Get-Date) -lt ($refresh)){
        Write-Logs "UPS token still valid"
        return
    }
    
    $cliPath = $cliPath = $env:USERPROFILE
    $fileName = "$cliPath\UPS.xml"
    try {
        $creds = Import-Clixml -Path $fileName
    } catch {
        Write-Logs "Couldn't get UPS creds"
    }
    
    $url = "https://onlinetools.ups.com/security/v1/oauth/token"
    
    $creds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $creds.UserName, $creds.GetNetworkCredential().password)))
    
    $body= @{
        "grant_type"= "client_credentials"
    }
    
    $headers= @{
        'Content-Type' = 'application/x-www-form-urlencoded'
        'x-merchant-id' = 'string'
        'Authorization' = 'Basic '+$creds
    }
    try{
        $result = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        $result.access_token | Out-File -FilePath C:\Users\MilesForet\Documents\Automations\UPS_Powershell\auth\bearer.txt
        Get-date ((Get-Date).AddHours(3).AddMinutes(59)) -UFormat "%m/%d/%y %I:%M %p" | Out-File -FilePath C:\Users\MilesForet\Documents\Automations\UPS_Powershell\auth\bearer_refresh.txt
        Write-Logs "Updated UPS Bearer Token"
    }catch{
        Write-Logs "*****ERROR: Failed to Udate UPS Bearer Token*****"
    }
}
