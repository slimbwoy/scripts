function global:send_email()
{


$CredUser = "alerts@computernetworksystems.us"
$CredPassword = "stlucianice"

$EmailFrom = "GAL_updater@graebelmoving.com"

$EmailTo = "melvin.felicien@graebelmoving.com,`
daniel.escaloni@graebelmoving.com,`
elroy@compnetsys.com,`
melvin.felicien@compnetsys.com,`
david.capone@graebelmoving.com,`
daniel.christenson@graebelmoving.com"

#$EmailTo = "melvin.felicien@graebelmoving.com" 



$Subject = "BRANCH TELEPHONE DIRECTORY"
$Body = @"
DEAR ALL,
      PLEASE FIND ATTACHED THE LATEST COPIES OF THE BRANCH DIRECTORIES GENERATED FROM THE GAL UPDATER
      THIS ALSO INCLUDES:
        - COMBINED FULL MIGRATED BRANCHES ONLY COPY (WITH MAIN NUMBERS)
        - COMBINED FULL COMPANY DIRECTORY (WITHOUT MAIN NUMBERS)

Regards,
Automated GAL Updater
"@

$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)

#GET FILES
$files = $wd + "\completed_directory"
$files = Get-ChildItem $files


Foreach($file in $files)
{
    $filepath = $file.FullName
   # write-host $filepath
    Write-Host “Attaching File :- ” $filepath
    $attachment = New-Object System.Net.Mail.Attachment("$filepath")
    $SMTPMessage.Attachments.Add($attachment)

}

#$SMTPServer = "smtpout.secureserver.net" 
#$SMTPServer = "mailer.graebelmoving.com"
$SMTPServer = "graebelmoving-com.mail.protection.outlook.com"
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25) 
$SMTPClient.EnableSsl = $false
#$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($CredUser, $CredPassword); 

$SMTPClient.Send($SMTPMessage)


}