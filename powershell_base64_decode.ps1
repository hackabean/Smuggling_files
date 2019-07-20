<# On linux: Use base64 file_name -w 0 #>

Function base64_decode

{

$base64 = "Place base64 encoded file here"

$fileContent = $base64
$fileContentBytes = [System.Convert]::FromBase64String($fileContent) 
[System.IO.File]::WriteAllBytes($env:USERPROFILE+"file\path.*",$fileContentBytes)
    
}

base64_decode
