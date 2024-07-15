Connect-PnPOnline -url "https://abs79.sharepoint.com/sites/ITLogsandAudits" -interactive

$query = @"
<View>
    <Query>
        <Where>
                    <Eq><FieldRef Name='Status'/><Value Type='Text'>Create Labels</Value></Eq>
        </Where>
    </Query>
</View>
"@

$sp_list = Get-PnPListItem -list "Equipment Prep" -Query $query
Connect-PnPOnline -url "https://abs79.sharepoint.com/sites/testing-miles" -interactive

foreach($item in $sp_list){

    #Add-PnPListItem -List "Create Labels" -Values @{"Person" = "Ebony Dickens"}
    $notes = $item["Notes"]

    #address
    $start = $notes.IndexOf("Home Address:")+14
    $end = $notes.IndexOf("Office:")
    $address = ($notes.ToString().SubString($start, $end-$start)).Trim()
    $add1 = ($address -split "\n")[0]
    
    $addline2 = ($address -split "\n")[1]

    if($addline2 -eq "" -or $null -eq $addline2){
        $city = ""
        $state = ""
        $zip = ""
    }else{
        $city = $addline2.ToString().SubString(0, $addline2.IndexOf(","))
        $start = $addline2.IndexOf(",")+1
        $end = $addline2.length-5
        $state = $addline2.ToString().SubString($start, $end-$start).Trim()
        $zip = $addline2.ToString().SubString($addline2.length-5, 5)
    }


    #phone

    $notes = $notes.Replace(" ","")
    $notes = $notes.Replace("-","")
    $notes = $notes.Replace("(","")
    $notes = $notes.Replace(")","")
    $start = $notes.IndexOf("Phone:")
    $phone = $notes.ToString().SubString($start+6, 10)

    $params = @{
        "Title" = $item["Title"]
        "PhoneNumber" = $phone
        "AddressLine1" = $add1
        "Address2" = ""
        "City" = $city
        "State_x0028_Abbr_x0029_" = $state
        "ZipCode" = $zip
        "ShipFrom" = $item["Assigned"]
        "EP_ID" = $item.Id
    }

    $params

    $i = Add-PnPListItem -List "Create Labels" -Values $params

    Clear-Variable -Name params
}