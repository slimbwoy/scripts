<#
    AUTOMATICALLY CONNECT TO OFFICE 365
    THIS IS A FUNCTION THAT CAN BE CALLED FROM ANY SCRIPT
#>

function o365_connect()
{

         #LOCAL ADMIN PASSWORD ENCRYPTION -------####
        #------------CRYPTOGRAPHY STUFF-----######
        $key = "212 61 164 87 170 45 200 228 33 2 213 187 28 186 206 169"
        $key = $key -replace " ",","
        $key = '(' + $key + ')'
        $key = (212,61,164,87,170,45,200,228,33,2,213,187,28,186,206,169)
        $encfile =  "$wd" + "\creds\enccreds.txt"
        $username = "melvin.felicien@graebelmoving.com"
        $EncryptedPW = Get-Content -Path $encfile
        $SecureString = ConvertTo-SecureString -String $EncryptedPW -Key $Key
        $global:o365cred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$SecureString
        #LOCAL ADMIN PASSWORD ENCRYPTION END ---###

        Connect-MsolService -Credential $o365cred

}

o365_connect
