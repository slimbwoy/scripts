<#

Export all the Sharepoint Online Site Collection Owners to a CSV File



#>
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$collowner = $wd + "\Exported_Owners\Collection_Owners.csv"
    Remove-Item $collowner -ErrorAction SilentlyContinue

$sendmail = $wd + "\send_mail.ps1"
    . $sendmail
    
Add-Type –Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type –Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$AdminUrl = "https://graebelmovingservices-admin.sharepoint.com/"
$UserName = "onedriveadmin@graebelmoving.com"
$Password = "Fu@uke7afene"
$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $SecurePassword

#Retrieve all site collection infos
Connect-SPOService -Url $AdminUrl -Credential $Credentials
#$sites = 
Get-SPOSite | Select-object url,owner | export-csv $collowner -NoTypeInformation
#write-host $sites


#send_email