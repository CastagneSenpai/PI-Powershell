<#
.SYNOPSIS
    Helper script to create scheduled task for script "PPR Centralized - PI DA Rotate Archives.ps1"

.DESCRIPTION
    Helper script that will create a scheduled task to run script named "PPR Centralized - PI DA Rotate Archives.ps1":
	- Both scripts have to be in same folder.
	- Task is created for system account to run it.
	- It runs every first day of the month at 01:00am.
	- Task is named $taskName.
	- Script uses schtasks.exe.

.NOTES
    Author: Anthony SABATHIER <anthony.sabathier@totalenergies.com>
    Version: 1.1
    Date: 2021-12-01
    Improvements: 
        - Task might be run as a service account
		- Move variables to parameters
		- Add logs
    Change Log:
		- v1.1: Updated script name.
#>
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
[string] $scriptPath = (Join-Path $PSScriptRoot 'PPR Centralized - PI DA Rotate Archives.ps1')
[string] $taskName = 'PI DA Archives Rotation (Test License)'
[string] $taskUser = 'NT AUTHORITY\SYSTEM'
# PowerShell.exe -ExecutionPolicy Bypass "& 'D:\Path\to\RotateArchives.ps1'"
[string]$command = "PowerShell.exe -ExecutionPolicy ByPass -File '$scriptPath'"

$taskParams = @(
    "/Create",
    "/RU", $taskUser,
    "/SC", "MONTHLY", 
    "/D",  "1",  # 1st Day of the month
    "/TN", $taskName, 
    "/TR", "$command",
    "/ST", "01:00",
    "/F",  #force
    "/RL", "HIGHEST" # run as admin
);

# supply the command arguments
schtasks.exe @taskParams

Write-Output $Error
[void][System.Console]::ReadKey($FALSE)