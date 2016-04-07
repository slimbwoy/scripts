<# 
    FETCH CONTACT INFORMATION FROM THE PBX AND UPLOAD TO SHAREPOINT ONLINE
    
    
#>

$host.ui.RawUI.WindowTitle = "GVL GAL / Directory Updater"
clear
$msg1 = READ-HOST -Prompt "`r`n`tSEND AN EMAIL CONTAINING THE DIRECOTORY FILES? (Yes/No) " 
$msg2 = read-host -Prompt "`r`n`tUPDATE GAL ON OFFICE 365? (Yes/No) "

$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
cd $wd
$connect_msol = $wd + "\connect-msol.ps1"
    . $connect_msol

$inc_email = $wd + "\send_mail.ps1"
    . $inc_email

    
$inc_json = $wd + "\includes\convertfromjson.ps1"
    #. $inc_json

$inc_disssl = $wd + "\includes\dis_ssl.ps1"
    . $inc_disssl

    Ignore-SelfSignedCerts

$testfile = $wd + "\testjson.txt"
    Remove-Item $testfile -Force -ErrorAction SilentlyContinue

    
$json_file = $wd + "\resources\contacts.json"
$csv = $wd + "\Contacts.csv"
    Remove-item $csv -ErrorAction SilentlyContinue -Force


$user = "gvldirectoryapi"
$pass = "Melvin12!"
$secpasswd = ConvertTo-SecureString $pass -AsPlainText -Force
$apicred = New-Object System.Management.Automation.PSCredential ($user, $secpasswd)


$apiurl = "https://pbx.graebelmoving.com:10000/api/user/contacts/company/"

Invoke-RestMethod -Uri $apiurl -Credential $apicred -Method get -OutFile $json_file


$read_json = (Get-Content $json_file -raw | ConvertFrom-Json ) 
$get_json = $read_json | select-object  LAST_NAME,FIRST_NAME,EMAIL,phone_numbers
foreach($entry in $get_json)
{

        $ext = $entry.phone_numbers.value -split " "
        $extension = $ext[0]
        $telephone = $ext[1]
        $ln = $entry.Last_name
        $fn = $entry.first_name
        $em = $entry.email

    $csvwrap = New-Object PSOBJECT -Property @{LASTNAME = $LN; FIRSTNAME = $FN; EMAIL = $EM; EXTENSION = $extension; TELEPHONE = $telephone}
    Export-Csv -InputObject $csvwrap $csv -NoTypeInformation -Append 


}


#EXTRACT THE DATA INTO USEABLE DATA--------------------------------


$CDirectory = $wd + "\Completed_Directory\FULL_Company_Directory.csv"
    Remove-Item $CDirectory -ErrorAction SilentlyContinue

$CDirectoryBad = $wd + "\Company_Directory_Bad.csv"
    Remove-Item $CDirectoryBad -ErrorAction SilentlyContinue

$CDhtml = $wd + "\Company_Directory_.html"
    Remove-Item $CDhtml -ErrorAction SilentlyContinue

$CDMig_full = $wd + "\Completed_Directory\FULL_Migrated_Directory.csv"
    Remove-item $CDMig_full -ErrorAction SilentlyContinue



Import-Csv $csv | ?{$_.email.length -gt 10} |?{$_.extension.length -gt 3} | select FIRSTNAME,LASTNAME,EMAIL,EXTENSION| Export-Csv -Path $CDirectory -NoTypeInformation
#Import-csv $csv | ConvertTo-Html -Body  $CDhtml
Import-Csv $csv |  ?{$_.email.length -lt 10} | select FIRSTNAME,LASTNAME,EMAIL,EXTENSION| Export-Csv -Path $CDirectoryBad -NoTypeInformation


##### MATCH EXTENSIONS TO MAIN OFFICE NUMBER

$mappingfile = $wd + "\resources\main_mapping.csv"
$read_map = import-csv $mappingfile

    foreach($office in $read_map)
       {
            #HEADER FIELDS : Office,MainNumber,startrange,endrange
            $officeloc = $office.office
            $mainn = $office.mainnumber -replace "-","."
            $srange = $office.startrange
            $erange = $office.endrange
            $bnum = $office.branchnumber

            $dir_by_branch = $wd + "\Completed_Directory\" + $officeloc.ToUpper() + ".csv"
                Remove-Item $dir_by_branch -ErrorAction SilentlyContinue -Force

            #CREATE A BREAKOUT FOR EACH BRANCH
            $match_offices = import-csv $CDirectory | Select-Object FIRSTNAME,LASTNAME,EMAIL,@{Name='MAIN_NUMBER';Expression={$mainn}},EXTENSION,@{Name='BRANCH_NUMBER';Expression={$bnum}} |
              ?{$_.extension -ge $srange -and  $_.extension -lt $erange} | 
              sort-object -Property FirstNAME  | 
              Export-Csv $dir_by_branch -NoTypeInformation

            #ADD / APPEND DATA TO MASTER DIRECTORY FILE
              $match_offices = import-csv $CDirectory | Select-Object FIRSTNAME,LASTNAME,EMAIL,@{Name='MAIN_NUMBER';Expression={$mainn}},EXTENSION,@{Name='BRANCH_NUMBER';Expression={$bnum}} |
              ?{$_.extension -ge $srange -and  $_.extension -lt $erange} | 
              sort-object -Property FirstNAME  | 
              Export-Csv $CDMig_full -NoTypeInformation -Append


            $matched_office = import-csv $dir_by_branch     
                   foreach($person in $matched_office)
                        { 
                             $fullname = "" + $person.firstname + " " + $person.lastname
                             $email = $person.email
                             $newnumber = ""+$person.main_number + " x" +$person.extension
                             $bnumber = $person.BRANCH_NUMBER
                             write-host "SETTING / APPLYING UPDATE FOR USER :$fullname : $email : $newnumber : $bnumber"
                            
                            #ONLY UPDATE OFFICE 365 GAL IF YES WAS SELECTED
                            if($msg2 -eq "YES")
                                {
                                    set-msoluser -UserPrincipalName $email -PhoneNumber "$newnumber" -Office "$bnumber"
                                }



                        }






       }


       #CLEANUP FULL FILE BY SORTING ALPHABETICALY
       (import-csv $CDMig_full | Sort-Object -Property FIRSTNAME ) | Export-Csv $CDMig_full -NoTypeInformation



       if($msg1 -eq "YES")
       {
            send_email

        }