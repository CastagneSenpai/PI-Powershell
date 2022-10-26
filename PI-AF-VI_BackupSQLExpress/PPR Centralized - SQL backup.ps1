<#
.SYNOPSIS
    Helper script to enforce PI AF & PI Vision SQL daily backup.

.DESCRIPTION
    Script meant to backup SQL Database of PI Suite installed locally in 3 steps for PIFD & PIVision databases:
    1. SQL Backup to $backupFolder\PIAF for PIFD & $backupFolder\PIVision for PIVision.
    2. Use 7-zip to compress bak file to minimize disk space usage.
    3. Retention policy applied to minimize disk space usage, keeping the file from 1st day of month.
    Also embbeds a way to install a scheduled task to run locally using localsytem daily at 1am.

.PARAMETER install
    Optional switch: Use installation mode for scheduled task.
    
.PARAMETER retentionDays
    Optional: Number of days defined by retention policy.
    Default: 7
        
.PARAMETER instance
    Optional: Target SQL Instance name.
    Default: Local SQLExpress instance.
    
.PARAMETER PIAFdatabase
    Optional: PI AF database name.
    Default: PIFD

.PARAMETER PIVIdatabase
    Optional: PI Vision database name.
    Default: PIVision

.PARAMETER logFile
    Optional: Full path to log file.
    Default: '.\logs\yyyy-MM_SQLBackup.log' under script folder

.PARAMETER backupFolder
    Optional: Path to backup folder.
    Default: F:\PIBackup

.NOTES
    Author: Anthony SABATHIER <anthony.sabathier@totalenergies.com>
    Version: 1.0
    Date: 2021-04-08
    Improvements: 
        - Logs compression after new month
        - Mail on error

.LINK
    - Sheduled Tasks CLI documenation: https://docs.microsoft.com/en-us/windows/win32/taskschd/schtask
    - SQL Server Backup cmdlet documentation page: https://docs.microsoft.com/en-us/powershell/module/sqlserver/backup-sqldatabase
    - 7-Zip home page: https://www.7-zip.org
    - PI EP SharePoint: https://totalworkplace.sharepoint.com/sites/PI/EP/
#>
param (
 [switch]$install,
 [int]$retentionDays = 7,
 [string]$instance = "$($ENV:COMPUTERNAME)\SQLExpress",
 [string]$PIAFdatabase = 'PIFD',
 [string]$PIVIdatabase = 'PIVision',
 [string]$logFile = (Join-Path $PSScriptRoot "logs\$(Get-Date -Format 'yyyy-MM')_SQLBackup.log"),
 [string]$backupFolder = 'F:\PIBackup'
)

Function Write-Log ($level, $message) {
    "$(Get-Date -UFormat '%Y-%m-%dT%T') $level $message" | Out-File -Append -FilePath $logFile
}

Function PIBackup-SQLDatabase ([string]$instance, [string]$database, [string]$backupFolder) {
    Write-Log "[DEBUG]" "Entering $($MyInvocation.MyCommand.Name)"
    Write-Log '[DEBUG]' ('Parameters: ' `
                    + 'Instance=' + $instance `
                    + ';Database=' + $database `
                    + ';BackupFolder=' + $backupFolder `
                    + ';logFile=' + $logFile + '.'
                    )
                    
    if (!(Test-Path $backupFolder)) {
        Write-Log '[DEBUG]' "'$backupFolder' does not exist, creating it."
        New-Item -ItemType Directory -Force -Path $backupFolder -ErrorAction Stop | Out-Null
    }
    [string]$backupFile = (Join-Path $backupFolder "$(Get-Date -Format 'yyyy-MM-dd')_$($instance.Replace('\','-'))_$($database).bak")
    Backup-SqlDatabase -ServerInstance "$instance" -Database "$database" -BackupAction database -BackupFile $backupFile -ErrorAction Stop
    Write-Log "[DEBUG]" "$database on $instance has been backed up successfully to $backupFile."
    return $backupFile
}

Function PIBackup-RotateBackups ([string]$backupFolder, [int]$retentionDays) {
    Write-Log "[DEBUG]" "Entering $($MyInvocation.MyCommand.Name)"
    Write-Log '[DEBUG]' "Cleaning files older than $retentionDays days in $($backupFolder). Also keeping first backup of each month."
    Get-ChildItem $backupFolder | Where-Object {$_.LastWriteTime -lt (get-date -Hour 0 -Minute 0 -Second 0).AddDays(-$retentionDays) -and (Get-Date -Date $_.LastWriteTime -Format 'dd') -ne '01'} | Remove-Item
    Write-Log '[DEBUG]' 'Cleaned up old backups successfully.'
}

# Apply compression to bak file.
Function PIBackup-CompressFile ([string]$backupFile) {
    Write-Log "[DEBUG]" "Entering $($MyInvocation.MyCommand.Name)"
    Write-Log "[DEBUG]" "File: $backupFile"
    # Check if 7zip 64-bits is available
    [string]$7zipPath = ""
    if (Test-Path "$($env:programfiles)\7-Zip\7z.exe") {
        $7zipPath = "$($env:programfiles)\7-Zip\7z.exe"
    } elseif (Test-Path "D:\Program Files\7-Zip\7z.exe") {
        $7zipPath = "D:\Program Files\7-Zip\7z.exe"
    } else {
        Throw "7z.exe not found at $($7zipPath). Please install 7-zip 64-bits version prior to run this script."
    }
    # a: archive | t7z: using 7z compression algorithm | mx9: ultra | -bsp0: no progress
    [string]$7zipArgs = "a","-t7z", "$backupFile.7z", "$backupFile","-mx9","-bsp0","-sdel"
    $7zipProcess = (Start-Process -FilePath $7zipPath -ArgumentList $7zipArgs -PassThru -Wait -ErrorAction Stop -NoNewWindow)
    if ($7zipProcess.ExitCode -gt 0 ) {
        Throw "Error while running 7-zip. Exit code: $($7zipProcess.ExitCode) (1 Warning, 2 Fatal, 7 CLI Error, 8 Not enough Memory, 255 User stopped). Please check '$outputFolder'"
    }
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Done."
}

Function Run() {
    try {
        Write-Log '[DEBUG]' "Entering 'PPR Centralized - SQL Backup' script."
        Write-Log '[DEBUG]' ('Parameters: ' `
                        + 'Instance=' + $instance `
                        + ';PIAFdatabase=' + $PIAFdatabase `
                        + ';PIVIdatabase=' + $PIVIdatabase `
                        + ';BackupFolder=' + $backupFolder `
                        + ';LogFile=' + $logFile + '.'
                        )
        [string]$PIAFBackupFile = PIBackup-SQLDatabase $instance $PIAFdatabase (Join-Path $backupFolder 'PIAF')
        [string]$PIVIBackupFile = PIBackup-SQLDatabase $instance $PIVIdatabase (Join-Path $backupFolder 'PIVision')
        PIBackup-CompressFile $PIAFBackupFile
        PIBackup-CompressFile $PIVIBackupFile
        PIBackup-RotateBackups (Join-Path $backupFolder 'PIAF') $retentionDays
        PIBackup-RotateBackups (Join-Path $backupFolder 'PIVision') $retentionDays
    } catch {
        Write-Log '[ERROR]' 'Error during backup execution.'
        Write-Log '[ERROR]' "Message: $($_.Exception.Message)"
    } finally {
        Write-Log '[DEBUG]' 'Script has completed. Bye :-)'
        Write-Log '[DEBUG]' '______________________________________________________'
    }
}

Function PIBackup-RegisterTask () {
    Write-Host "$(Get-Date -UFormat '%Y-%m-%dT%T') [DEBUG] Entering 'PPR Centralized - SQL Backup' script in installation mode"
    $taskParams = @(
        "/Create",
        "/RU", 'NT AUTHORITY\SYSTEM',
        "/SC", "DAILY",
        "/TN", 'PI PPR Centralized - SQL Backup',
        "/TR", "PowerShell.exe -ExecutionPolicy ByPass -File '$PSCommandPath'",
        "/ST", "01:00",
        "/F"#,  #force
        #"/RL", "HIGHEST" # run as admin
    );
    
    Write-Host "$(Get-Date -UFormat '%Y-%m-%dT%T') [DEBUG] Running command 'schtasks.exe $taskParams'"
    schtasks.exe @taskParams
    Write-Host $Error
    pause
}

Function Main () {
    # In case target logs file folder does not exist.
    if (!(Test-Path (Split-Path -parent $logFile))) {
        New-Item -ItemType Directory -Force -Path (Split-Path -parent $logFile) -ErrorAction Stop
    }

    # Install task vs Run
    if ($install) {
        # Run As Admin if needed
        if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -install" -Verb RunAs; exit }
        PIBackup-RegisterTask
    } else {
        Run
    }
}

Main
