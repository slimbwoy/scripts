#
#  SYNOPSIS :
 
#  This script download the MBAMclient from COMPNETSYS REPO
 
#  Written by the Team at comptuer Network systems : wwww.compnetsys.com
 
#  Mckin S
#  Melvin F
 
#
 $global:logfile = $temp_path + "\gfi_scep_deploy.log"  
  
######-----VERIFY THAT WE ARE OKAY TO INSTALL ------####
### cHECK THAT THE MACHINE IS ON THE NA DOMAIN & THAT SCEP IS NOT YET INSTALLED ON THE MACHINE---#####
function TestScep(){
    
    $temp_path = $env:SystemRoot + "\temp"
    $domain = $ENV:USERDOMAIN 
    $workingdomain = "na"
    
    $reg_key = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Security Client"
    $test_key = Test-Path $reg_key


   
    $runtime = Get-Date -Format yyyy_M_dd_h_m
    
    if(($test_key -eq "True" -AND $domain -eq $workingdomain))
        { 
            $scep_version = Get-ItemProperty $reg_key | select -ExpandProperty DisplayVersion
            clear
            echo  "`r`n`t$runtime : SOFTWARE VERSION $SCEP_VERSION ALREADY INSTALLED -- QUITTING`r`n"
            echo  "`r`n`t$runtime : SOFTWARE VERSION $SCEP_VERSION ALREADY INSTALLED -- QUITTING`r`n" | Out-File $logfile -Append
            echo "`r`n ------- END ------- END--------------`r`n" |  Out-File $logfile -Append
            exit
            }        
            else
            {
                clear
                echo  "`r`n`t$runtime : PROCEEDING WITH INSTALL - PREVIOUS VERSION NOT DETECTED`r`n"
                echo  "`r`n`t$runtime : PROCEEDING WITH INSTALL - PREVIOUS VERSION NOT DETECTED`r`n" | Out-File $logfile -Append
              
            }


}

#call the function
    TestScep

####-----END VERIFY THAT SOFTWARE IS / ISNT INSTALLED -----########


#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:dst_folder = "C:\ProgramData\scepinstall"
#we need to specify the actual destination file name. 
$global:dst_file = "C:\ProgramData\scepinstall\scep.exe"

 $install_exe = "C:\ProgramData\scepinstall\Scep\SCEPInstall.exe"

#The folder to place the downloaded file in, must exist before the .net.webclient will successfully download the file.
#Remove-item $dst_folder -Force -Recurse -ErrorAction SilentlyContinue

if(!(Test-path $global:dst_folder)) { New-Item $global:dst_folder -ItemType Directory }
$global:src_file = "https://s3-us-west-2.amazonaws.com/gms-software-dist/Scep.exe"

      
  
 
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
 echo "`r`n`tSTARTING SOFTWARE DOWNLOAD" |  Out-File $logfile -Append

    download $src_file $dst_file

    Start-Process $dst_file -Wait

 echo "`r`n`tCOMPLETED SOFTWARE DOWNLOAD" |  Out-File $logfile -Append



#EXTRACT THE ZIP FILE


 # ---------------INSTALL CLIENT ---------------------------------------##################


    echo "`r`n`tSTARTING SOFTWARE INSTALL" |  Out-File $logfile -Append
   
        Start-Process $install_exe -ArgumentList "/s /q"

    echo "`r`n`tSOFTWARE INSTALL COMPLETED" |  Out-File $logfile -Append



 #####------------------------------------###########
 

   echo "`r`n ------- END ------- END--------------`r`n" |  Out-File $logfile -Append