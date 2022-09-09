Import-Module -Name OSDProgress -Force

Write-Host  -ForegroundColor Cyan "Sono qui!"

Start-Sleep -Seconds 5

Watch-OSDCloudProvisioning {
    Write-Host -ForegroundColor Cyan "Hey this script running an OSD Cloud ZTI Deployment while displaying a MahApps.Metro progress window"

    #Start OSDCloud ZTI
    Update-OSDProgress -Text "Running OSDCloud PreAction stuff..." # output to UI
    Write-Host  -ForegroundColor Cyan "Running OSDCloud PreAction stuff..." # output to console
    Start-Sleep -Seconds 5
    Update-OSDProgress -Text " " # hide first text

    Start-OSDCloud -OSVersion "Windows 11" -OSBuild 21H2 -OSLanguage en-us -OSEdition Pro -ZTI

    #Anything I want  can go right here and I can change it at any time since it is in the Cloud!!!!!
    Update-OSDProgress -Text "Running OSDCloud PostAction stuff..."
    Write-Host  -ForegroundColor Cyan "Running OSDCloud PostAction stuff..."
    Start-Sleep -Seconds 5
    Update-OSDProgress -Text " " # hide first text

    # lets throw an error, just for fun
    #Update-OSDProgress -DisplayError "Custom error message, pls unlock screen!"

    #Restart from WinPE
    Update-OSDProgress -Text "Reboot in 20 seconds"
    Start-Sleep -Seconds 20
    wpeutil reboot
} -Window -Style Win11
