$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$encfile =  "$wd" + "\enccreds.txt"
Function RandomKey {
     $RKey = @()
     For ($i=1; $i -le 16; $i++) {
     [Byte]$RByte = Get-Random -Minimum 0 -Maximum 256
     $RKey += $RByte
     }
     $RKey
}
$Key = RandomKey
$ScureString = ConvertTo-SecureString -String "Welcome@1!" -AsPlainText -Force
$EncryptedPW = ConvertFrom-SecureString -SecureString $ScureString -Key $Key
Set-Content -Path "$encfile" -Value $EncryptedPW
