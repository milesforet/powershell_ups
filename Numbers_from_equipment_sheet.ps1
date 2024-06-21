Connect-PnPOnline -url "https://abs79.sharepoint.com/sites/ITLogsandAudits" -interactive

$shipped = @"
    <View>
        <Query>
            <Where>
                <Or>
                    <Eq><FieldRef Name='Status'/><Value Type='Text'>Shipped</Value></Eq>
                    <Eq><FieldRef Name='Status'/><Value Type='Text'>Shipped to Office</Value></Eq>
                </Or>
            </Where>
        </Query>
    </View>
"@

try {

    $sp_list = Get-PnPListItem -list "Equipment Prep" -Query $shipped

    <#
    Iterate through items in list. Check to make sure shipping info is not empty
    Split the multi-line string into an array. Each index being a line from the string iterate through the array. 
    Find the tracking number in the line. Searches for index of "1Z" then uses that index to substring the line for the tracking number
    Adds any valid tracking numbers to the UPS Tracking Sharepoint List
    #>
    if(!$sp_list){
        Write-Host "No new items to ship"

    }else{
        foreach($item in $sp_list){
    
            if(!$item["ShippingInfo"]){
        
                Write-Host $item["Title"]"- shipping Info is empty. Could not add to UPS Tracking list"
            
            }else{
                
                $tracking_nums = ($item["ShippingInfo"]).Split(“`n”)
        
                foreach($line in $tracking_nums){    
        
                    $index = $line.IndexOf("1Z")
        
                    if($index -eq -1){
        
                        Write-Host "Line does not contain valid tracking number"
        
                    }else{
        
                        try {
                            $tracking_num = $line.ToString().SubString($index, 18)
                            #$i = Add-PnPListItem -List "UPS Tracking" -Values @{"Title"=$tracking_num}
                            Write-Host $tracking_num" added to UPS Tracking List"
                        }
        
                        catch {
                            Write-Host -f red "Encountered Error:"$_.Exception.Message
                        }
                    }
                }
            }
        }
    }
}
catch {
    Write-host -f red "Encountered Error:"$_.Exception.Message
}