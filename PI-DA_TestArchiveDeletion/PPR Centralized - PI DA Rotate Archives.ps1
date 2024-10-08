<#
.SYNOPSIS
    Helper script to enforce PI Data Archive test server's 3 rolling months retention policy. (rounded to 95days default)

.DESCRIPTION
    Based on OSIsoft's MoveOldArchives.ps1 script.
    Helper script to enforce PI Data Archive test server's 3 rolling months retention policy.
    Enables one to delete all archives on a server that are older than provided number of days (default is 95 days).

.PARAMETER PIDAServerName
    Optional: Name of the PIDA Server we want to backup.
    Default: Current hostname ($ENV:COMPUTERNAME)

.PARAMETER retentionDays
    Optional: Number of days defined by retention policy.
    Default: 95

.PARAMETER logDir
    Optional: Directory to put logs into. Not used if logFile is set manually.
    Default: $PSScriptRoot\logs

.PARAMETER logFile
    Optional: Full path to log file.
    Default: 'yyyy-MM_RotateArchives.log' in logDir.

.PARAMETER errorMailCC
    Optional: Parameter to defined new recipients of alerts in case of errors
    Default: HQ support team.

.NOTES
    Author: Anthony SABATHIER <anthony.sabathier@totalenergies.com>
    Version: 1.1
    Date: 2021-12-01
    Improvements: 
        - Logs compression after new month
    Change Log:
        - V1.1: Updated mails for TotalEnergies, fixed typo line.101
#>
param(
	[string] $PIDAServerName = $ENV:COMPUTERNAME,
	[int] $RetentionPolicyDays = 95,
    [string] $logDir = (Join-Path $PSScriptRoot 'logs'),
    [string] $logFile = (Join-Path $logDir "$(Get-Date -Format yyyy-MM)_PIDA_RotateArchives.log"),
	[string] $errorMailCC = 'anthony.sabathier@totalenergies.com'#'EP DSI-SUPPORT-PI <ep.dsi-support-pi@totalenergies.com>'
)

# Helper Function: Log writer into $logFile
Function Write-Log ([string]$level, [string]$message) {
    "$(Get-Date -UFormat '%Y-%m-%dT%T') $level $message" | Out-File -Append -FilePath $logFile
    "$(Get-Date -UFormat '%Y-%m-%dT%T') $level $message" | Write-Host
}

# Send mail in case of error, Name provided.
Function Send-MailError ([string]$body) {
    $encodingMail = [System.Text.Encoding]::UTF8
    [string]$to  = 'anthony.sabathier@totalenergies.com'#"ep.dsi-support-pi@totalenergies.com"
    [string]$from = "pi-da-archive-rotation@mail01.totalenergies.com"
    [string]$smtpServer = "EMEAMAICLI-EL01.main.glb.corp.local"
    [string]$subject = "[PIDA Archives Rotation] Failure - $($ENV:COMPUTERNAME) - $(Get-Date -format yyyy-MM-dd)"
    try {
        Send-MailMessage -to $to -From $from -Cc $errorMailCC -Subject $subject -SmtpServer $smtpServer -BodyAsHtml $body -Encoding $encodingMail -ErrorAction Stop
    } catch {
        Write-Log "[ERROR]" "$($_.Exception.ToString())"
    }
}

Function RotateArchives ([string] $PIDAServerName, [int] $RetentionPolicyDays) {
    Write-Log "[DEBUG]" "Rotating archives for $PIDAServerName older than $RetentionPolicyDays days"
    # Get the PI server connection object, exit if there is an error retrieving the server
    $connection = Connect-PIDataArchive -PIDataArchiveMachineName $PIDAServerName -ErrorAction Stop

    [Version] $v395 = "3.4.395"
    if ($connection.ServerVersion -gt $v395) {
       $archives = Get-PIArchiveFileInfo -Connection $connection -ArchiveSet 0 -ErrorAction Stop
    } else {
       $archives = Get-PIArchiveFileInfo -Connection $connection -ErrorAction Stop
    }

    # Iterate through each archive and test to see if the archive should be moved
    foreach ($archive in $archives) {
       if ($archive.Index -ne 0 -and                      		  # Make sure that we aren't looking at the primary archive
           $archive.EndTime -ne $null -and                        # Verify that the archive has a start time
           $archive.EndTime -ne [System.DateTime]::MinValue -and  # Verify that the archive start time is set
           $archive.EndTime -lt (Get-Date).AddDays(-$RetentionPolicyDays) # Check to see if the archive start time is older than the time specified
           ) {
          Write-Log "[DEBUG]" "Unregistering archive: $($archive.Path)"
          # Unregister the archive
          Unregister-PIArchive -Name $archive.Path -Connection $connection -ErrorAction Stop
          Write-Log "[DEBUG]" "Removing archive: $($archive.Path)"
          # Remove the archive
          Remove-Item $archive.Path -Force -ErrorAction Stop
       }
    }
}


try {
    if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop }
    RotateArchives $PIDAServerName $RetentionPolicyDays
} catch {
    Write-Log "[ERROR]" "Error rotating archive. Detail: $($_.Exception.ToString())"
    Send-MailError "PI DA Archives Rotation Script has failed for $(Get-Date -format yyyy-MM-dd) run. Please check logs on server. Error: $($_.Exception.ToString())"
}
