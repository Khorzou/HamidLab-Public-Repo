# SYNOPSIS
# Create Intune_Multi_administrator_approval_Notification task scheduler


# DEMO
# YouTube video link → https://www.youtube.com/@chandermanipandey8763

# NOTES
# Version:         V1.0  
# Author:          Chander Mani Pandey 
# Creation Date:   24 March 2026

# Find the author on: 
 
# YouTube:         https://www.youtube.com/@chandermanipandey8763  
# Twitter:         https://twitter.com/Mani_CMPandey  
# LinkedIn:        https://www.linkedin.com/in/chandermanipandey  
# BlueSky:         https://bsky.app/profile/chandermanipandey.bsky.social
# GitHub:          https://github.com/ChanderManiPandey2022
#------------------------------ User Input Section Start --------------------------------------------------

$TaskName = "Intune_Multi_administrator_approval_Notification"
$Description = "Intune_Multi-administrator approval"
$ScriptPath = "C:\Intune_Report_Automation\IntuneMultiAdminApprovalMailNotifications.ps1"

# Start time (first run)
$StartTime = Get-Date "09:30AM"

#------------------------------ User Input Section End -----------------------------------------------------

# Create task action
$action = New-ScheduledTaskAction -Execute "Powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""

# Create 24 hourly triggers
$triggers = @()

for ($i = 0; $i -lt 24; $i++) {
    $runTime = $StartTime.AddHours($i)
    $triggers += New-ScheduledTaskTrigger -Daily -At $runTime
}

# Register scheduled task
Register-ScheduledTask -User SYSTEM `
    -Action $action `
    -Trigger $triggers `
    -TaskName $TaskName `
    -Description $Description `
    -Force

Write-Host "$TaskName successfully created with 24 hourly triggers." -ForegroundColor Green