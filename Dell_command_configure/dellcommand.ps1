 
#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:dst_folder = "C:\ProgramData\dellcommand"
 
#The folder to place the downloaded file in, must exist before the .net.webclient will successfully download the file.
if(!(Test-path $global:dst_folder)) { New-Item $global:dst_folder -ItemType Directory }
 
$global:src_file = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/Command_Configure.msi"
#we need to specify the actual destination file name. 
$global:dst_file = "C:\ProgramData\dellcommand\Dellcommand.msi"
 
 
 
            function global:download($src,$dst)
                {
 
                   WRITE-HOST "`r`nDOWNLOADING LATEST MBAM CLIENT FILES $src ----->`r`n"
                   IF(TEST-PATH  $dst){Remove-Item $dst -ErrorAction SilentlyContinue}
 
                   #Invoke-WebRequest $src -OutFile $dsst
 
                   $DOWNLOAD = New-Object System.Net.WebClient
                   $DOWNLOAD= $DOWNLOAD.DownloadFile($src,$dst)
 
 
                   if(Test-Path $dst) { WRITE-HOST "`tDOWNLOAD COMPLETED FOR $src"}
                }
 
$global:hostname = $env:computername
 
 #Initiate Download
 download $src_file $dst_file
 
 # Install Dell Command 
 #$logfile = $env:SystemRoot + "\temp\dellcommand.log"
 Start-Process msiexec "/i $dst_file /qb ALLUSERS=1 REBOOT=ReallySuppress"
 
 
 
 
function global:send_email()
{
 
$CredUser = "alerts@computernetworksystems.us"
$CredPassword = "stlucianice"
 
$EmailFrom = "Mbam-client@graebelmoving.com"
$EmailTo = "salesforceinstalls@compnetsys.com" 
$Subject = "$subject"
$Body = "$msg"
 
$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
#$attachment = New-Object System.Net.Mail.Attachment($LOGFILE)
#$SMTPMessage.Attachments.Add($attachment)
#if(Test-Path $SFILOG)
#{
#$attachment2 = New-Object System.Net.Mail.Attachment($SFILOG)
#$SMTPMessage.Attachments.Add($attachment2)
#}
 
 
 
$SMTPServer = "smtpout.secureserver.net" 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 80) 
$SMTPClient.EnableSsl = $false
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($CredUser, $CredPassword); 
 
$SMTPClient.Send($SMTPMessage)
 
 
}
 
clear
 
#$testfile1 = $dst + "\"+$file1 
 
if(Test-Path $dst_file)
{
    $global:subject = "SUCCEEDED : Dell Command Configure  $HOSTNAME"
    send_email
 
}
 
else
{
 
 
$global:subject = "FAILED : Dell Command Configure $HOSTNAME"
    send_email
 
 }