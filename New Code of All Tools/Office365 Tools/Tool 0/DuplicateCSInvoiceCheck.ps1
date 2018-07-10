# Short Description: Identify Duplicate records from coversheet
#
# Long Description:
# Check duplicate records from coversheet which processed = false
# Iteration of each records which not processed checking again with same library with Invoice ID
# Duplicate record exist if record found greater than 1
# Duplicate record exist if ID of both record NOT matched
# Duplicate record exist if Invoice Date of both record is matched
# Duplicate record exist if Invoice Amount of both record is matched
# Duplicate record exist if Vondor No. of both record is matched
# C1=If any record passed from above condition then check if records have processed status is False
# If C1 then Updating List 2 record's status processed = Duplicate
# Otherwise Updating List 1 record's status processed = Duplicate
#
# @author: Mahendra Vala 
# @license  http://www.gkblabs.com/ GKBLabs License
# @version  Powershell V1, ISAPI V1
# @date: 18/05/2018

$Logfile = "C:\Dayton Rogers Log History\DuplicateCoversheet-"+(Get-Date).ToString("yyyMMdd") +".log"
$CoversheetFile = "C:\Dayton Rogers Log History\DuplicateCoversheet-"+(Get-Date).ToShortDateString()+".csv"
$ProcessedStatusArray = @("True", "Error", "Archive");


$logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO ==================== Started Duplicate Invoice Processing ===================="
Add-Content $Logfile -Value $logContent

$logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Adding Path"
Add-Content $Logfile -Value $logContent
# Adding Path
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"  
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"  

$logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Seting configuration variables"
Add-Content $Logfile -Value $logContent

# Seting configuration variables
$siteURL = "https://gkblabs1.sharepoint.com/sites/DaytonRogers/"
$userId = "adi@gkblabs.com"

#$siteURL = "https://daytonrogers.sharepoint.com/sites/DaytonRogersInvoiceProcessing/"
#$userId = "spadmin@daytonrogers.com"

$pwd = ConvertTo-SecureString "HmN68nl4zL@123" -AsPlainText -Force 
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)

$logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Creating sharepoint objects"
Add-Content $Logfile -Value $logContent

$ctx.credentials = $creds
try{
    # Setting CamlQuery Object
    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Setting CamlQuery Object"
    Add-Content $Logfile -Value $logContent
    $lists = $ctx.web.Lists
    $list = $lists.GetByTitle("Coversheet")
    # CamlQuery Object
    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO CamlQuery Object"
    Add-Content $Logfile -Value $logContent
    $ViewXml = "<View><Query><Where><Eq><FieldRef Name='processed'/><Value Type='Choice'>False</Value></Eq></Where></Query></View>";
    $cquery2 = New-Object Microsoft.SharePoint.Client.CamlQuery
    $cquery2.ViewXml=$ViewXml

    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO "+$ViewXml
    Add-Content $Logfile -Value $logContent
    Write-Host "List 1 Query " $ViewXml

    $listItems = $list.GetItems($cquery2)

    $ctx.load($listItems)  
    
    Write-Host "Execute Query"
    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Execute Query"
    Add-Content $Logfile -Value $logContent
    $ctx.executeQuery()

    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Unprocessed invoice count: "+ $listItems.Count
    Add-Content $Logfile -Value $logContent
    Write-Host "Unprocessed invoice count: " $listItems.Count

    $list2 = $lists.GetByTitle("Coversheet")
    $i = 0

    $coversheetRecords = "Invoice Number,Invoice Amount,Invoice Date,Vendor ID,Vendor Name,Title`r`n"

    if ($listItems.Count -gt 0) {
        foreach($listItem in $listItems)  
        {
            $processedStatus = "False"
            if ($i -gt 0) {
                # CamlQuery 3
                $ViewXml3 = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>"+$listItem['ID']+"</Value></Eq></Where></Query></View>"
                $cquery = New-Object Microsoft.SharePoint.Client.CamlQuery

                $cquery.ViewXml=$ViewXml3

                $listItems3 = $list2.GetItems($cquery)

                $ctx.load($listItems3)

                # Executing Query
                $ctx.executeQuery()
                $processedStatus = $listItems3[0]["processed"]
            }

            if ($processedStatus -eq "Duplicate") { continue }

            Write-Host "Invoice # - " $listItem["Invoice_x0020__x0023_"] "Title - " $listItem["Title"]  
            $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Invoice # - " + $listItem["Invoice_x0020__x0023_"] + "Title - " + $listItem["Title"] 
            Add-Content $Logfile -Value $logContent

            Write-Host "ID: " $listItem['ID'] "Invoice Date - " $listItem['Invoice_x0020_Date'] "Invoice Amt: " $listItem['Invoice_x0020_Amt'] " Vendor: " $listItem['Vendor']
            $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO ID: " + $listItem['ID'] + "Invoice Date - " + $listItem['Invoice_x0020_Date'] + "Invoice Amt: " + $listItem['Invoice_x0020_Amt'] + " Vendor: " + $listItem['Vendor']
            Add-Content $Logfile -Value $logContent

            # CamlQuery 2
            $ViewXml2 = "<View><Query><Where><And><Eq><FieldRef Name='Invoice_x0020__x0023_'/><Value Type='Text'>"+$listItem['Invoice_x0020__x0023_']+"</Value></Eq><Neq><FieldRef Name='ID' /><Value Type='Counter'>"+$listItem['ID']+"</Value></Neq></And></Where></Query></View>"
            $cquery = New-Object Microsoft.SharePoint.Client.CamlQuery

            $cquery.ViewXml=$ViewXml2

            Write-Host "CamlQuery 2" $ViewXml2
            $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO CamlQuery 2" + $ViewXml2
            Add-Content $Logfile -Value $logContent

            $listItems2 = $list2.GetItems($cquery)

            $ctx.load($listItems2)

            # Executing Query
            $ctx.executeQuery()  

            Write-Host "Total " $listItems2.Count " Record(s) found for Invoice #" $listItem['Invoice_x0020__x0023_']
            $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Total " + $listItems2.Count + " Record(s) found for Invoice #" + $listItem['Invoice_x0020__x0023_']
            Add-Content $Logfile -Value $logContent

            if ($listItems2.Count -ge 1)
            {

                foreach($listItem2 in $listItems2)  
                {
                    # Record information
                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Record information"
                    Add-Content $Logfile -Value $logContent
                    Write-Host "ID: " $listItem2['ID'] "Invoice Date - " $listItem2['Invoice_x0020_Date'] "Invoice Amt: " $listItem2['Invoice_x0020_Amt'] " Vendor: " $listItem2['Vendor'] " Status: " $listItem2['processed']
                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO ID: " + $listItem2['ID'] + "Invoice Date - " + $listItem2['Invoice_x0020_Date'] + "Invoice Amt: " + $listItem2['Invoice_x0020_Amt'] + " Vendor: " + $listItem2['Vendor'] + " Status: " + $listItem2['processed']
                    Add-Content $Logfile -Value $logContent

                    # List 1 Invoice Date & List 2 Invoice Date matched!
                    if ($listItem['Invoice_x0020_Date'] -eq $listItem2['Invoice_x0020_Date']) {
                        Write-Host "# List 1 Invoice Date & List 2 Invoice Date matched!"
                        $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO List 1 Invoice Date & List 2 Invoice Date matched!"
                        Add-Content $Logfile -Value $logContent
                    
                        # List 1 Invoice Amount & List 2 Invoice Amount matched!
                        if ($listItem['Invoice_x0020_Amt'] -eq $listItem2['Invoice_x0020_Amt']) {
                            Write-Host "# List 1 Invoice Amount & List 2 Invoice Amount matched!"
                            $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO List 1 Invoice Amount & List 2 Invoice Amount matched!"
                            Add-Content $Logfile -Value $logContent

                            if ($listItem['Vendor'] -eq $listItem2['Vendor']) {
                                Write-Host "# List 1 Vendor ID & List 2 Vendor ID matched!"
                                $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO List 1 Vendor ID & List 2 Vendor ID matched!"
                                Add-Content $Logfile -Value $logContent

                                Write-Host "Duplicate found!... " -foregroundcolor black -backgroundcolor green
                                $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Duplicate found!... "
                                Add-Content $Logfile -Value $logContent

                                if ($listItem2['processed'] -eq "False") {
                                    Write-Host "Updating ID" $listItem2['ID'] " Invoice #" $listItem2['Invoice_x0020__x0023_']
                                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Updating ID" + $listItem2['ID'] + " Invoice #" + $listItem2['Invoice_x0020__x0023_']
                                    Add-Content $Logfile -Value $logContent
                                    Write-Host "Updating List 2 !..."
                                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Updating List 2 !..."
                                    Add-Content $Logfile -Value $logContent

                                    $listItem2['processed'] = "Duplicate"
                                    $listItem2.Update();
                                    # Pushing update to server!
                                    Write-Host "Pushing update to server!" -foregroundcolor black -backgroundcolor green
                                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Pushing update to server!"
                                    Add-Content $Logfile -Value $logContent
                                    $ctx.ExecuteQuery()
                                    
                                    $coversheetRecords += "'"+$listItem2['Invoice_x0020__x0023_']+"','"+$listItem2['Invoice_x0020_Amt']+"','"+$listItem2['Invoice_x0020_Date']+"','"+$listItem2['Vendor']+"','"+$listItem2['Vendor_x0020_Name']+"'`r`n"

                                    $i++
                                }
                                else
                                {
                                    
                                    if ($listItem2['processed'] -in $ProcessedStatusArray) {
                                        Write-Host "Updating ID" $listItem['ID'] " Invoice #" $listItem['Invoice_x0020__x0023_']
                                        $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Updating ID" + $listItem2['ID'] + " Invoice #" + $listItem2['Invoice_x0020__x0023_']
                                        Add-Content $Logfile -Value $logContent
                                        Write-Host "Updating List 1 !..."
                                        $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Updating List 1 !..."
                                        Add-Content $Logfile -Value $logContent
                                        
                                        $listItem['processed'] = "Duplicate"
                                        $listItem.Update();
                                        # Pushing update to server!
                                        Write-Host "Pushing update to server!" -foregroundcolor black -backgroundcolor green
                                        $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Pushing update to server!"
                                        Add-Content $Logfile -Value $logContent
                                        $ctx.ExecuteQuery()
                                        $coversheetRecords += "'"+$listItem2['Invoice_x0020__x0023_']+"','"+$listItem2['Invoice_x0020_Amt']+"','"+$listItem2['Invoice_x0020_Date']+"','"+$listItem2['Vendor']+"','"+$listItem2['Vendor_x0020_Name']+"'`r`n"

                                    }
                                    else
                                    {
                                        Write-Host "Nothing updated as everything already set!" -foregroundcolor black -backgroundcolor red
                                        $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Nothing updated as everything already set!"
                                        Add-Content $Logfile -Value $logContent
                                    }
                                    $i++
                                }
                                # Updating List 2 for logic purpose
                                Write-Host "Updating List 2 Modified Date for logic purpose"
                                $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Updating List 2 Modified Date for logic purpose"
                                Add-Content $Logfile -Value $logContent
                                $listItem2['Modified'] = [System.DateTime]::Now
                                $ctx.ExecuteQuery()
                            }
                        }
                    }
                }
            }
            else
            {
                Write-Host "Record not found in Coversheet: $($listname2) to duplicate" -foregroundcolor black -backgroundcolor red
                $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Record not found in Coversheet: $($listname2) to duplicate"
                Add-Content $Logfile -Value $logContent
            }
        }
    }
    else 
    {
        Write-Host "Record not found in Coversheet: $($listname) to duplicate" -foregroundcolor black -backgroundcolor red
        $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Record not found in Coversheet: $($listname) to duplicate"
        Add-Content $Logfile -Value $logContent
    }
    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO ==================== Completed Duplicate Invoice Processing ===================="
    Add-Content $Logfile -Value $logContent
    Add-Content $CoversheetFile -Value $coversheetRecords


}
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " ERROR $($_.Exception.Message)"
    Add-Content $Logfile -Value $logContent
}