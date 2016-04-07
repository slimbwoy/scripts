$allusers = $ENV:allusersprofile
$scanlocation = "$allusers" + "\Microsoft\Microsoft Antimalware\Scans\History\Service\DetectionHistory"
$count_files = (Get-Childitem $scanlocation -recurse | ?{ ! $_.PSIsContainer } | measure-object).Count
if($count_files -gt "0")
    {
        echo "VIRUS DETECTION - PLEASE CHECK HOST"
        echo "exit 1"
        exit 1
        } 
        else 
        {
            exit 0
        }