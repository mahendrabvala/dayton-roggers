# Mahendra Vala 
# 12/04/2017
# Recursively crawls the directory tree and converts tiff files to pdf files
# compresses the images to a smaller file size.

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"  
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"  
  
$siteURL = "https://gkblabs1.sharepoint.com/sites/DaytonRogers/"
$userId = "adi@gkblabs.com"

$siteURL = "https://daytonrogers.sharepoint.com/sites/DaytonRogersInvoiceProcessing/"
$userId = "spadmin@daytonrogers.com"

#$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
$pwd = ConvertTo-SecureString "TIbyw3oJ" -AsPlainText -Force 

$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds  

try{  
    $lists = $ctx.web.Lists  
    $list = $lists.GetByTitle("Coversheet")
    $ViewXml = "<View><Query><Where><Eq><FieldRef Name='processed'/><Value Type='Choice'>False</Value></Eq></Where></Query></View>";
    $cquery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $cquery.ViewXml=$ViewXml
        
    Write-Host $ViewXml
        
    $listItems = $list.GetItems($cquery)

    $ctx.load($listItems)  
      
    $ctx.executeQuery()  

    Write-Host $listItems.Count

    foreach($listItem in $listItems)  
    {  
        Write-Host "Invoice # - " $listItem["Invoice_x0020__x0023_"] "Title - " $listItem["Title"]  
        $listItem['processed'] = "True"
        $listItem.Update();
        $ctx.ExecuteQuery()
        Write-Host "Exist kare che" -foregroundcolor black -backgroundcolor green
        
    }
}  
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
}