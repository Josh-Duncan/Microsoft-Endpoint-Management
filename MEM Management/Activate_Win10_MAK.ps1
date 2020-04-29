 
$ProductKey = "Change this to your product key"

$License = Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" | where { $_.PartialProductKey } | select Description, LicenseStatus 

if ($License.LicenseStatus -ne "1") 
{

    Write-Host $License.Description"is not activated, activating..."     

    try{
        Start-Process -FilePath "cmd.exe" -ArgumentList "-cmd /k cls && cscript.exe slmgr.vbs /ipk $ProductKey && exit" -wait | Out-Null
        write-host "Activated"
        }
    catch{write-error "ACTIVATION FAILED"}
}

elseif ($License.LicenseStatus -eq "1"){Write-Host "OS already activated"} 
