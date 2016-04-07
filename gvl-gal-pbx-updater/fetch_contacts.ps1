<# 
    FETCH CONTACT INFORMATION FROM THE PBX AND UPLOAD TO SHAREPOINT ONLINE
    
    
#>




$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
cd $wd
$connect_msol = $wd + "\connect-msol.ps1"
    . $connect_msol

$Global:exempt_file = $wd + "\exemptions.txt"

    
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


$CDirectory = $wd + "\Company_Directory.csv"
    Remove-Item $CDirectory -ErrorAction SilentlyContinue

$CDirectoryBad = $wd + "\Company_Directory_Bad.csv"
    Remove-Item $CDirectoryBad -ErrorAction SilentlyContinue

$CDhtml = $wd + "\Company_Directory_.html"
    Remove-Item $CDhtml -ErrorAction SilentlyContinue

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

            $dir_by_branch = $wd + "\Completed_Directory\" + $officeloc.ToUpper() + ".csv"
                Remove-Item $dir_by_branch -ErrorAction SilentlyContinue -Force

            $match_offices = import-csv $CDirectory | Select-Object FIRSTNAME,LASTNAME,EMAIL,@{Name='MAIN_NUMBER';Expression={$mainn}},EXTENSION   |  ?{$_.extension -ge $srange -and  $_.extension -lt $erange} | sort-object -Property FirstNAME  | Export-Csv $dir_by_branch -NoTypeInformation
            $matched_office = import-csv $dir_by_branch     
                   foreach($person in $matched_office)
                        { 
                             $exempt = $null
                             $fullname = "" + $person.firstname + " " + $person.lastname
                             $email = $person.email
                             $newnumber = ""+$person.main_number + " x" +$person.extension
                             write-host "SETTING / APPLYING UPDATE FOR USER :$fullname : $email : $newnumber"
                            

                             #SOME USERS HAVE MORE THAN 1 NUMBER. OMIT THESE USERS WHEN THEY ASK FOR A MANUAL UPDATE
                             $exempt = (select-string $exempt_file -pattern $email| measure-object).count

                             if($exempt -ne 0)
                             {
                                $email = $email.toupper()
                                write-host "`r`n`tEXEMPTION FOUND FOR $EMAIL" -ForegroundColor Green
                                
                                }
                                else
                                {
                                    #UPDATE RECORD IN OFFICE 365
                                    $checkuser = Get-MsolUser -UserPrincipalName $email -ErrorAction SilentlyContinue
                                        if($checkuser)
                                         {
                                            set-msoluser -UserPrincipalName $email -PhoneNumber "$newnumber"
                                            }
                                    }

                                    
                        }
                        


       }


