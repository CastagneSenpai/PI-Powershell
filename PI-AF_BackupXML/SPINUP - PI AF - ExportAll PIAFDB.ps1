<#
.SYNOPSIS
    Export all AF Databases of provided server into provided folder with compression as part of backup. (includes retention policy) 

.DESCRIPTION
    This script is a backup tool provided as part of PI backup strategy.
    It enables backuping of provided AFServerName into outputFolder using compression.
    If backup runs ok, it also deletes files in outputFolder if they are older than retention days (policy) provided.
    If server can send e-mails, can trigger e-mails to PI HQ support team in case of fatal issue.

.PARAMETER AFServerName
    Name of the PIAF Server we want to backup.

.PARAMETER outputFolder
    Path to backup folder.

.PARAMETER retentionDays
    Optional: Number of days defined by retention policy.
    Default: 7

.PARAMETER folderSuffix
    Optional: Suffix for backup folder name creation.
    Default: '_<AFServerName>_Backup_XML'

.PARAMETER logFile
    Optional: Path to log file.
    Default: 'yyyy-MM_ExportAll_PIAFDB.log' in script folder's "Logs" subfolder

.PARAMETER errorMailCC
    Optional: Parameter to defined new recipients of alerts in case of errors
    Default: None.

.NOTES
    Authors: 
        - Jonathan BARON <j.baron@cgi.com>
        - Anthony SABATHIER <anthony.sabathier@totalenergies.com>
    Version: 2.3
    Date: 2021-12-06
    Improvements: 
        - Fallback to standard windows compression if no 7zip installed
        - Remove adaptations for PowerShell v2.0 after SPIN'UP project
        - Add size check so see if next backup will fail
		- Remove Event Frame backup (just need a backfilling to regenerate them if needed)
    Changes:
        - v2.2 - Worked on retention actions to avoid issue with relative path
        - v2.3 - Changed mail domain to totalenergies.com, Add patch for nigeria on database name, Changed default log folder and create it if missing
#>
param (
    [Parameter(Mandatory=$True)][string]$AFServerName,
    [Parameter(Mandatory=$True)][string]$outputFolder,
    [int]$retentionDays = 7,
	[string]$folderSuffix = "_$($AFServerName)_Backup_XML",
    [string]$logFile = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "Logs\$(Get-Date -Format yyyy-MM)_ExportAll_PIAFDB.log"),
    [string[]]$errorMailCC = 'ep.dsi-support-pi@totalenergies.com'
)

# Helper Function: Log writer into $logFile
Function Write-Log ([string]$level, [string]$message) {
    "$(Get-Date -UFormat '%Y-%m-%dT%T') $level $message" | Out-File -Append -FilePath $logFile
    "$(Get-Date -UFormat '%Y-%m-%dT%T') $level $message" | Write-Host
}

# Returns a folder name with date stamp and suffix provided.
Function GetFolder-Timestamped ([string]$folderRoot, [string]$folderSuffix) {
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Generating timestamped folder in '$folderRoot'."
    [string]$Date = (Get-Date -format yyyy-MM-dd)
    [string]$timestampedFolder = (Join-Path $folderRoot "$($Date)$($folderSuffix)")
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Attempting to create '$timestampedFolder'"
	$newFolder = (New-Item -ItemType Directory -Force -Path "$timestampedFolder" -ErrorAction Stop)
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Created '$($newFolder.fullname)'."
    return ($newFolder.FullName)
}

# Export ALL AF Databases for AF Server in exportFolder
Function ExportXML-AFServer([string]$AFServerName, [string]$exportFolder) {
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Exporting all bases for '$AFServerName' to '$exportFolder'"
    [Reflection.Assembly]::LoadWithPartialName("OSIsoft.AFSDK") | Out-Null
    [OSIsoft.AF.PISystems]$AFServers = New-Object OSIsoft.AF.PISystems
    [OSIsoft.AF.PISystem]$AFServer = $AFServers["$AFServerName"]
    [System.EventHandler]$EventHandler

    foreach ($AFDatabase in $AFServer.Databases) {
        Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Doing '$($AFDatabase.name)'"
		[string]$AFDatabaseName_Replaced = ($AFDatabase.Name).replace('/','_')
        # 49=1+16+32 => https://techsupport.osisoft.com/Documentation/PI-AF-SDK/html/T_OSIsoft_AF_PIExportMode.htm
        $AFServer.ExportXML($AFDatabase, 49, (join-path $exportFolder "$($AFDatabaseName_Replaced).xml"),$null,$null, $EventHandler)
	}
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Done."
}

# Apply compression to folder.
Function Compress-Folder ([string]$outputFolder) {
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Compressing '$outputFolder'."
    # Check if 7zip 64-bits is available
    [string]$7zipPath = ""
    if (Test-Path "$($env:programfiles)\7-Zip\7z.exe") {
        $7zipPath = "$($env:programfiles)\7-Zip\7z.exe"
    } elseif (Test-Path "D:\Program Files\7-Zip\7z.exe") {
        $7zipPath = "D:\Program Files\7-Zip\7z.exe"
    } else {
        Throw "7z.exe not found at $($env:programfiles)\7-Zip\7z.exe nor at D:\Program Files\7-Zip\7z.exe. Please install 7-zip 64-bits version prior to run this script."
    }
    # a: archive | t7z: using 7z | mx9: ultra | -bsp0: no progress | sdel: source deletion
    [string]$7zipArgs = "a","-t7z ""$outputFolder.7z"" ""$outputFolder\""","-mx9","-bsp0","-sdel"
    $7zipProcess = (Start-Process -FilePath $7zipPath -ArgumentList $7zipArgs -PassThru -Wait -ErrorAction Stop -NoNewWindow)
    if ($7zipProcess.ExitCode -gt 0 ) {
        Throw "Error while running 7-zip. Exit code: $($7zipProcess.ExitCode) (1 Warning, 2 Fatal, 7 CLI Error, 8 Not enough Memory, 255 User stopped). Please check '$outputFolder'"
    }
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Done."
}

# Delete files and folders older than provided number of days in provided folder path.
Function Rotate-Backups ([string]$targetFolder, [int]$retentionDays) {
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Removing items older than $retentionDays day(s) from '$targetFolder'."
    
	#$files = (Get-ChildItem -Path $targetFolder -Recurse | Where-Object { $_.PSIsContainer -And $_.LastWriteTime -lt (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-$retentionDays) })
    
	$RetentionDate = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-$retentionDays)
    $files = (Get-ChildItem -Path $targetFolder -Recurse | Where-Object {$_.LastWriteTime -lt $RetentionDate })
	
	foreach ($file in $files) {
	Write-Log "[DEBUG]" "Removed file: $($file.fullName)"
        try {
			Write-Log "[DEBUG]" "Removed file: $($file.fullName)"
            $file | Remove-Item -ErrorAction Stop
            Write-Log "[DEBUG]" "Removed file: $($file.fullName)"
        } catch {
            Write-Log "[ERROR]" "Error removing file: $($file.fullName)."
            Write-Log "[ERROR]"  $_.Exception.ToString()
        }
    }
    $emptyFolders = (Get-ChildItem -Path $targetFolder -Recurse | Where-Object { $_.PSIsContainer -eq $true -And (Get-ChildItem -Path $_.FullName).Count -eq 0 })
    foreach ($emptyFolder in $emptyFolders) {
        try {
            $emptyFolder | Remove-Item -Force -ErrorAction Stop
            Write-Log "[DEBUG]" "Removed folder: $($emptyFolder.fullName)"
        } catch {
            Write-Log "[ERROR]" "Error removing folder: $($emptyFolder.fullName)."
            Write-Log "[ERROR]"  $_.Exception.ToString()
        }
    }
    Write-Log "[DEBUG]" "$($MyInvocation.MyCommand.Name): Done."
}

# Send mail in case of error, Name provided.
Function Send-MailError ([string]$body) {
    $encodingMail = [System.Text.Encoding]::UTF8
    [string]$to  = "ep.dsi-support-pi@totalenergies.com"
    [string]$from = "pi-af-xml-backup@mail01.totalenergies.com"
    [string]$smtpServer = "EMEAMAICLI-EL01.main.glb.corp.local"
    [string]$subject = "[PIAF XML Backup] Failure - $($ENV:COMPUTERNAME) - $AFServerName - $(Get-Date -format yyyy-MM-dd)"
    try {
        Send-MailMessage -to $to -From $from -Cc $errorMailCC -Subject $subject -SmtpServer $smtpServer -BodyAsHtml $body -Encoding $encodingMail -ErrorAction Stop
    } catch {
        Write-Log "[ERROR]" "$($_.Exception.ToString())"
    }
}

# Wrapping script actions into Main function
Function Main () {
    try {
        # Step 0 - In case log file parent folder does not exists.
        if (!(Test-Path (Split-Path -Path $logFile -Parent))) {
            New-Item -ItemType Directory -Path (Split-Path -Path $logFile -Parent) -Force -ErrorAction Stop
        }
        # Step 1 - generate the output Folder
        [string]$exportFolder = (GetFolder-Timestamped $outputFolder $folderSuffix)
        # Step 2 - backup all af datbases to XML inside folder from step 1
        ExportXML-AFServer $AFServerName $exportFolder
        # Step 3 - archive and compress folder from step 1 and files from step 2
        Compress-Folder $exportFolder
        # Step 4 - everything is ok: remove old backups
        Rotate-Backups $outputFolder $retentionDays
    } catch {
        Write-Log "[ERROR]" "XML Export Script has failed with error: $($_.Exception.ToString())"
        Send-MailError "AF XML Export Script has failed for $(Get-Date -format yyyy-MM-dd) run. Please check logs. Error: $($_.Exception.ToString())"
        Write-Error "AF XML Export Script has failed. Check '$logFile' for detailled error." -ErrorAction Stop
    }
}

Main
