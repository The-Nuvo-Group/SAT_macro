function getXLSTARTpath($pcUser){
    $pre = "C:\Users\"
    $post = "\AppData\Roaming\Microsoft\Excel\XLSTART"

    return $pre+$pcUser+$post
}

$xlsbFiles = (Get-ChildItem -Path ".\" -Filter "*.XLSB").FullName 

$pcUser = $([System.Environment]::UserName)

$XLSTARTPath =  getXLSTARTpath($pcUser)


Write-Host
Write-Host "------------------------------------"
Write-Host
Write-Host "Setting up SAT"
Write-Host
Write-Host "User: " $pcUser
Write-Host "Targets:"
foreach($i in (Get-ChildItem -Path ".\" -Filter "*.XLSB").Name)
{
Write-Host "`t-> "$i
}
Write-Host "Destination: " $XLSTARTPath
Write-Host
Write-Host "------------------------------------"
Write-Host

Foreach($i in $xlsbFiles){
    Move-Item -Path $i -Destination $XLSTARTPath
}

dir $XLSTARTPath

Write-Host
Write-Host "------------------------------------"
Write-Host
Write-Host "ALL DONE HERE!"
Read-Host -Prompt "Press ENTER to finishe"
Write-Host 
Write-Host "------------------------------------"