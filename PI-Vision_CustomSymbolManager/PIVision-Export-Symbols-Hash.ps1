<#
.SYNOPSIS
Helper script to pull all custom symbol files on PI Vision servers.

.DESCRIPTION
This script pulls all custom symbol files hashes for given PI Vision server list. 
It generates one timestamped CSV file. Naming pattern is: yyyy-MM-ddTHH-mm-ss_PIVision_CustomSymbols.csv

.PARAMETER inputFile
Path to CSV file containing at least one column "ServerName" (PI Vision server) and "PIVisionExtPath" (PI Vision server 'ext' folder path UNC format).

.PARAMETER outputFolder
[OPTIONAL] Folder for output CSV file. To be set for custom target folder.
[DEFAULT] PSScriptRoot (Folder where script is located).

.PARAMETER logFile
[OPTIONAL] Path to log file. To be set for custom target file.
[DEFAULT] PSScriptRoot\yyyy-MM-dd_PIVision-Export-Symbols-Hash.log (File in folder where script is located).

.NOTES
    FILE:    PIVision-Export-Symbols-Hash.ps1
    AUTHOR:  Anthony SABATHIER <anthony.sabathier@totalenergies.com>
    VERSION: 1.1

#> 
param (
    # To be used for multi server mode, if run against localhost only, no need to provide.
    [Parameter(Mandatory=$true)][string]$inputFile,
    # Custom output folder if needed.
    [string]$outputFolder = $PSScriptRoot,
    # Custom log file if needed.
    [string]$logFile = (Join-Path $PSScriptRoot "$(Get-Date -Format 'yyyy-MM-ddTHH-mm-ss')_PIVision-Export-Symbols-Hash.log")
    
)

##############################################################################
# Utility fonction for logging purposes, quite fancy in this script
##############################################################################
Function Write-Log ([string]$level, [string]$message) {
    try {
        [string]$logString = "$(Get-Date -UFormat '%Y-%m-%dT%T') $level $message"
        Out-File -InputObject $logString -FilePath $logFile -Append 
        [string]$color = "white"
        if ($level -like "*ERR*") { $color = "red" }
        if ($level -like "*WARN*") {$color = "yellow" }
        if ($level -like "*SUCESS*") {$color = "green" }
        Write-Host $logString -ForegroundColor $color
    } catch {
        Write-Log '[ERROR]' $_.Exception.Message
    }
}

##############################################################################
# Function with main routine for a single server
##############################################################################
Function Get-RemoteFilesHash ([string]$serverName, [string]$localRootPath) {
    Write-Log "DEBUG" "Entering Get-PIVision-SymbolFiles. (serverName='$serverName',localRootPath='$localRootPath')"
    [string]$remoteRootPath = (Join-Path "\\$serverName\" $localRootPath.Replace(':','$'))
    if (!(Test-Path $remoteRootPath)) {
        Write-Log "ERROR" "Not able to list files in '$remoteRootPath'."
        return
    }
    # We retrieve all files recursively (subfolders) and get their SHA512 hash
    $remoteFiles = Get-ChildItem $remoteRootPath -Recurse | % { Get-FileHash $_.FullName -Algorithm SHA512 }
    # $remoteFiles
    return $remoteFiles 
}

##############################################################################
# Main function with loop logic over several servers
##############################################################################
Function Main() {
    import-module (Join-Path $PSScriptRoot '..\lib\logs.psm1') 

    Write-Log "DEBUG" "Entering PIVision-Export-Symbols-Hash script."
    # Verifying that we have an input file available.
    if (!(Test-Path $inputFile)) {
        Write-Log "ERROR" "No input file available at '$inputFile'."
        return 
    }
    # Making output folder if not exists
    if (!(Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder
    }
    [string]$outputFile = (Join-Path $outputFolder ("$(Get-Date -Format 'yyyy-MM-dd')_PIVision_CustomSymbols.csv"))
    # Looping over input file rows
    $PIVisionServers = Import-Csv $inputFile
    foreach ($PIVisionServer in $PIVisionServers) {
        # Adding files hash list into the input table.
        Add-Member -InputObject $PIVisionServer `
            -NotePropertyName "HashFilesTable" `
            -NotePropertyValue (Get-RemoteFilesHash $PIVisionServer.ServerName $PIVisionServer.PIVisionExtPath)
    }
    # Export to csv
    $PIVisionServers.HashFilesTable | Export-Csv -UseCulture -Encoding UTF8 -NoTypeInformation $outputFile

    # All good, bye bye.
    Write-Log "DEBUG" "Exiting PIVision-Export-Symbols-Hash script."
}

##############################################################################
# Running main
Main
# End of script
##############################################################################
