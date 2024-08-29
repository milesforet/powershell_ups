$creds = Get-AbsCred -credName "service-equip@abskids.net"
Connect-PnPOnline -url "https://abs79.sharepoint.com/sites/ITLogsandAudits" -Credentials $creds

$query = @"
<View>
    <Query>
        <Where>
            <And>
                <Eq><FieldRef Name='Status'/><Value Type='Text'>Incoming</Value></Eq>
                <Neq><FieldRef Name='AddressInfo'/><Value Type='Text'>True</Value></Neq>
            </And>
        </Where>
    </Query>
</View>
"@

$sp_list = Get-PnPListItem -list "Equipment Prep" -Query $query

if($sp_list.length -eq 0){
    Write-Logs "No incoming items"
    exit
}

foreach($item in $sp_list){

    #Add-PnPListItem -List "Create Labels" -Values @{"Person" = "Ebony Dickens"}
    $notes = $item["Notes"]

    #address
    $start = $notes.IndexOf("Home Address:")+14
    $end = $notes.IndexOf("Office:")
    if($end -eq -1){
        Write-Logs -msg "*****ERROR: $($item["Title"]) - Couldn't get address info*****"
        continue
    }

    $address = ($notes.ToString().SubString($start, $end-$start)).Trim()
    $add1 = ($address -split "\n")[0]
    $add2 = ($address -split "\n")[1]

    if($add1 -match "unit"){
        $add2Index = $add1.IndexOf("unit", [StringComparison]"CurrentCultureIgnoreCase")
    }elseif ($add1-match "apt") {
        $add2Index = $add1.IndexOf("apt", [StringComparison]"CurrentCultureIgnoreCase")
    }elseif($add1 -match "#"){
        $add2Index = $add1.IndexOf("#", [StringComparison]"CurrentCultureIgnoreCase")
    }
    
    $addline1 = ""
    $addline2 = ""
    
    if($add2Index){
        $addline1 = $add1.Substring(0, $add2Index)
        $addline2 = $add1.Substring($add2Index, ($add1.Length)-$add2Index)
    }else{
        $addline1 = $add1
    }


    if($add2 -eq "" -or $null -eq $add2){
        $city = ""
        $state = ""
        $zip = ""
    }else{
        $city = $add2.ToString().SubString(0, $add2.IndexOf(","))
        $start = $add2.IndexOf(",")+1
        $end = $add2.length-5
        $state = $add2.ToString().SubString($start, $end-$start).Trim()
        $zip = $add2.ToString().SubString($add2.length-5, 5)
    }

    #phone
    $notes = $notes.Replace(" ","")
    $notes = $notes.Replace("-","")
    $notes = $notes.Replace("(","")
    $notes = $notes.Replace(")","")
    $start = $notes.IndexOf("Phone:")
    $phone = $notes.ToString().SubString($start+6, 10)
    
    $sup_query = @"
<View>
    <Query>
        <Where>
                <Eq><FieldRef Name='Username'/><Value Type='Text'>$($item["ABSUser"])</Value></Eq>
        </Where>
    </Query>
</View>
"@

    #supervisor
    Connect-PnPOnline -url "https://abs79.sharepoint.com/sites/ITLogsandAudits" -Credentials $creds
    $supervisor = Get-PnPListItem -List "Onboarding Queue" -Query $sup_query

    $bill_to = ""

    switch -Wildcard ($supervisor["Role"]){
        "*Scheduling*" {$bill_to = "Scheduling $($supervisor["ReportsToText"])"}
        "*Biller*" {$bill_to = "Billing $($supervisor["ReportsToText"])"}
        "*Revenue Cycle*" {$bill_to = "Billing $($supervisor["ReportsToText"])"}
        "*Psychologist*" {$bill_to = "Psych $($supervisor["ReportsToText"])"}
        "*Intake*" {$bill_to = "Intake $($supervisor["ReportsToText"])"}
        "*IT Support*" {$bill_to = "IT $($supervisor["ReportsToText"])"}
        "*Data Analyst*" {$bill_to = "Data/Reporting $($supervisor["ReportsToText"])"}
        "*HR*" {$bill_to = "HR $($supervisor["ReportsToText"])"}
        "*Accounts Payable*" {$bill_to = "AP $($supervisor["ReportsToText"])"}
        "*Application Support*" {$bill_to = "IT Systems $($supervisor["ReportsToText"])"}
        "*System Support*" {$bill_to = "IT Systems  $($supervisor["ReportsToText"])"}
        "*Authorizations" {$bill_to = "Benefits and Auth $($supervisor["ReportsToText"])"}
        "*Recruiter*" {$bill_to = "Recruiting $($supervisor["ReportsToText"])"}
        "*Learning" {$bill_to = "Learning $($supervisor["ReportsToText"])"}
        "*Business Development*" {$bill_to = "Business Dev $($supervisor["ReportsToText"])"}
        "*Operations Manager*" {$bill_to = "Clinical Ops $($supervisor["ReportsToText"])"}
    }


    $params = @{
        "PhoneNumber" = $phone
        "AddressLine1" = $addline1.trim()
        "AddressLine2" = $addline2
        "City" = $city
        "State_x0028_Abbr_x0029_" = $state
        "ZipCode" = $zip
        "AddressInfo" = "True"
        "BillTo" = $bill_to
    }

    $params

    try{
        $i = Set-PnPListItem -Identity $item.Id -List "Equipment Prep" -Values $params
    }catch{
        Write-Logs "ERROR $($item["Title"]) - $_"
    }

    $params = @{
        SupervisorContact = $supervisor["ReportsToText"]
    }

    #check for display names that are different than legal names
    switch($params.SupervisorContact){
        "Nicole Young" {$params.SupervisorContact = "Dr. Nicole Young"}
        "Vincent Pope" {$params.SupervisorContact = "Vince Pope"}
        "Jeffrey Hinckley" {$params.SupervisorContact = "Jeff Hinckley"}
        "Stephen Yeager Jr" {$params.SupervisorContact = "Steve Yeager Jr"}
    }


    try{
        $i = Set-PnPListItem -Identity $item.Id -List "Equipment Prep" -Values $params
    }catch{
        Write-Logs "ERROR $($item["Title"]) - $_"
    }

    Write-Logs -msg "$($item["Title"]) - Address info has been updated"

    Clear-Variable -Name params
}