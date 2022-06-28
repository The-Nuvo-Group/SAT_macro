function getXLSTARTpath($pcUser){
    $pre = "C:\Users\"
    $post = "\AppData\Roaming\Microsoft\Excel\XLSTART"

    return $pre+$pcUser+$post
}

$pcUser = $([System.Environment]::UserName)

$XLSTARTPath =  getXLSTARTpath($pcUser)

$currentLocation = [string](Get-Location)

$fileName = "ANNEXA1- V1 ALPHA.XLSB"

$target = $currentLocation+"\"+$fileName

Write-Host
Write-Host "------------------------------------"
Write-Host
Write-Host "Setting up SAT"
Write-Host
Write-Host "User: " $pcUser
Write-Host "Target: '$fileName' "
Write-Host "Destination: " $XLSTARTPath
Write-Host
Write-Host "------------------------------------"
Write-Host

Move-Item -Path $target -Destination $XLSTARTPath

dir $XLSTARTPath

Write-Host
Write-Host "------------------------------------"
Write-Host
Write-Host "ALL DONE HERE!"
Read-Host -Prompt "Press ENTER to finishe"
Write-Host 
Write-Host "------------------------------------"