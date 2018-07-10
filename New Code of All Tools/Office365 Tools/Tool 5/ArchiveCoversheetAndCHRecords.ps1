# Short Description: Mark the records status to Archive from coversheet & Coversheet Header
#
# Long Description:
# Query to Coversheet Header library with condtiion Divisional Approval = Divisional Approval Approved
# Iteration of each coversheet header records which Divisional Approval Approved
# Check record has Accounting Approval: Accounting Approval Approved
# Check record modified date is older than 30 days
# Query to Coversheet for the Invoice ID which false true in above condtions
# Iteration of each coversheet records & update them to Processed = Archive
# Update Coversheet Header records to Processed = Archive
#
# @author: Mahendra Vala 
# @license  http://www.gkblabs.com/ GKBLabs License
# @version  Powershell V1, ISAPI V1
# @date: 25/06/2018


$Logfile = "C:\Dayton Rogers Log History\ArchivalTool-"+(Get-Date).ToShortDateString()+".log"
$CoversheetHeaderFile = "C:\Dayton Rogers Log History\CoversheetHeader-"+(Get-Date).ToShortDateString()+".csv"
$CoversheetFile = "C:\Dayton Rogers Log History\Coversheet-"+(Get-Date).ToShortDateString()+".csv"

$logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO ==================== Started Invoice Archival ===================="
Add-Content $Logfile -Value $logContent

$logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Adding Path"
Add-Content $Logfile -Value $logContent
# Adding Path
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"  
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"  

$logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Seting configuration variables"
Add-Content $Logfile -Value $logContent

# Seting configuration variables
$siteURL = "https://gkblabs1.sharepoint.com/sites/DaytonRogers/"
$userId = "adi@gkblabs.com"
$pwd = "HmN68nl4zL@123"  

$siteURL = "https://daytonrogers.sharepoint.com/sites/DaytonRogersInvoiceProcessing/"
$userId = "spadmin@daytonrogers.com"
# $pwd = "TIbyw3oJ"  

$pwd = Read-Host -Prompt "Enter password" -AsSecureString  


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

    $listCoverSheetHeader = $lists.GetByTitle("CoverSheetHeader")
    $listCoversheet = $lists.GetByTitle("Coversheet")

    # CamlQuery Object
    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO CamlQuery Object"
    Add-Content $Logfile -Value $logContent
    $ViewXml = "<View><Query><Where><And><Eq><FieldRef Name='Divisional_x0020_Approval'/><Value Type='Choice'>Divisional Approval Approved</Value></Eq><Eq><FieldRef Name='Accounting_x0020_Approval' /><Value Type='Choice'>Accounting Approval Approved</Value></Eq></And></Where></Query></View>";
    $cquery2 = New-Object Microsoft.SharePoint.Client.CamlQuery
    $cquery2.ViewXml=$ViewXml

    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO "+$ViewXml
    Add-Content $Logfile -Value $logContent
    Write-Host "List 1 Query " $ViewXml

    $listCoverSheetHeaderItems = $listCoverSheetHeader.GetItems($cquery2)

    $ctx.load($listCoverSheetHeaderItems)  
    
    Write-Host "Execute Query"
    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Execute Query"
    Add-Content $Logfile -Value $logContent
    $ctx.executeQuery()

    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Divisional Approval Approved invoices count: "+ $listCoverSheetHeaderItems.Count
    Add-Content $Logfile -Value $logContent
    Write-Host "Divisional Approval Approved invoices count: " $listCoverSheetHeaderItems.Count
    $coversheetHeaderRecords = "Invoice Number,Invoice Amount,Invoice Date,Vendor ID`r`n"
    $coversheetRecords = "Invoice Number,Invoice Amount,Invoice Date,Vendor ID`r`n"
    if ($listCoverSheetHeaderItems.Count -gt 0) {
        foreach($listItem in $listCoverSheetHeaderItems) {
            $DateBefore30Days = (get-date).AddDays(-30)

            $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Applying filter of Modified is 30 days older "
            Add-Content $Logfile -Value $logContent
            Write-Host " Applying filter of Modified is 30 days older "
            if ($DateBefore30Days -ge $listItem['Modified']) {

                $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Applying filter of Accounting Approval: Accounting Approval Approved "
                Add-Content $Logfile -Value $logContent
                Write-Host " Applying filter of Accounting Approval: Accounting Approval Approved "

                if ($listItem["Accounting_x0020_Approval"] -eq 'Accounting Approval Approved') {

                    $coversheetHeaderRecords += $listItem["Title"] +","+ $listItem["Invoice_x0020_Amount"]+","+ $listItem["Invoice_x0020_Date"]+","+ $listItem["Vendor_x0020_Number"]+"`r`n"

                    Write-Host "Vendor Invoice Number: " $listItem["Title"]  " Modified: " $listItem["Modified"] " Divisional Approval Status: " $listItem["Divisional_x0020_Approval"] 

                    # $ViewXml = "<View><Query><Where><And><Eq><FieldRef Name='Invoice_x0020__x0023_' /><Value Type='Text'>"+$listItem["Title"]+"</Value></Eq><Eq><FieldRef Name='processed' /><Value Type='Choice'>False</Value></Eq></And></Where></Query></View>";
                    $ViewXml = "<View><Query><Where><Eq><FieldRef Name='Invoice_x0020__x0023_' /><Value Type='Text'>"+$listItem["Title"]+"</Value></Eq></Where></Query></View>";
                    $cquery = New-Object Microsoft.SharePoint.Client.CamlQuery
                    $cquery.ViewXml=$ViewXml

                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO "+$ViewXml
                    Add-Content $Logfile -Value $logContent
                    Write-Host "List 2 Query " $ViewXml

                    $listCoversheetItems = $listCoversheet.GetItems($cquery)

                    $ctx.load($listCoversheetItems)  
    
                    Write-Host "Execute Query"
                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Execute Query"
                    Add-Content $Logfile -Value $logContent
                    $ctx.executeQuery()
                    
                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Divisional & Accounting Approval Approved Coversheet invoices count: "+ $listCoversheetItems.Count
                    Add-Content $Logfile -Value $logContent
                    Write-Host "Divisional & Accounting Approval Approved Coversheet invoices count: " $listCoversheetItems.Count

                    if ($listCoversheetItems.Count -gt 0) {
                        foreach ($listCoversheetItem in $listCoversheetItems) {
                            if ($listItem['Invoice_x0020_Date'] -eq $listCoversheetItem['Invoice_x0020_Date']) {                    
                                if ($listItem['Invoice_x0020_Amount'] -eq $listCoversheetItem['Invoice_x0020_Amt']) {
                                    if ($listItem['Vendor_x0020_Number'] -eq $listCoversheetItem['Vendor']) {
                                        $coversheetRecords += $listCoversheetItem["Invoice_x0020__x0023_"] +","+ $listCoversheetItem["Invoice_x0020_Amt"]+","+ $listCoversheetItem["Invoice_x0020_Date"]+","+ $listCoversheetItem["Vendor"]+"`r`n"

                                        Write-Host " Coversheet: Updating status processed = 'Archive "

                                        $listCoversheetItem['processed'] = "Archive"
                                        $listCoversheetItem.Update();
                                        # Pushing update to server!
                                        Write-Host "Pushing update to server!" -foregroundcolor black -backgroundcolor green
                                        $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Pushing update to server!"
                                        Add-Content $Logfile -Value $logContent
                                        $ctx.ExecuteQuery()
                                    }
                                }
                            }
                        }
                    }
                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO CoversheetHeader: Updating status processed = 'Archive' "
                    Add-Content $Logfile -Value $logContent
                    Write-Host " CoversheetHeader: Updating status processed = 'Archive' "

                    $listItem['processed'] = "Archive"
                    $listItem.Update();
                    # Pushing update to server!
                    Write-Host "Pushing update to server!" -foregroundcolor black -backgroundcolor green
                    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Pushing update to server!"
                    Add-Content $Logfile -Value $logContent
                    $ctx.ExecuteQuery()
                }
            } else  {
                Write-Host  Write-Host "Vendor Invoice Number: " $listItem["Title"]  " is not Account Approval Approved" -foregroundcolor black -backgroundcolor red
                $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " INFO Record not found in Coversheet: $($listname) to duplicate"
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
                    
    Add-Content $CoversheetHeaderFile -Value $coversheetHeaderRecords
    Add-Content $CoversheetFile -Value $coversheetRecords

}
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
    $logContent = ( Get-Date).ToShortDateString() + " " + (Get-Date).ToLongTimeString() + " ERROR $($_.Exception.Message)"
    Add-Content $Logfile -Value $logContent
}