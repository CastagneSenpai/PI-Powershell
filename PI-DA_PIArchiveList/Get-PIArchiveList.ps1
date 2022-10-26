<#
.SYNOPSIS
List all archives (Corrupted or not) of a PI Server. 


.DESCRIPTION
This script is meant to help list archives for an entire PI Server. 

 

.PARAMETER corruptOnly
[Optional] Switch to specify we want to list corrupted archives



.PARAMETER outputFile
[Optional] Select the outputfile directory
 
 

.NOTES
Title: Get-PIArchiveList
Author: Melvin CARRERE <melvin.carrere@external.total.com>
Version: 0.1
#>

Param(
[Parameter(HelpMessage="Do you want to filter the archive corrupted files ? (y/n)")][String]$corruptOnly = "n",
[Parameter(HelpMessage="Insert your outputfolder")][String]$outputFile = ".\output\ArchiveList_$timestamp.csv",
[Parameter(HelpMessage="Insert the log path folder")][String]$v_logPathfile = ".\logs\logs_$timestamp.txt"
)

import-module (Join-Path $PSScriptRoot '..\Lib\Logs.psm1')
#import-module .\Lib\Logs.psm1

Clear-Host

#Initialization of variables, logs and inputs.
$affiliates = Import-Csv ".\input\affiliate.csv"
$timestamp=get-date -uFormat "%m%d%Y%H%M"
$v_logPathfile = ".\logs\logs_$timestamp.txt"

Write-Log -v_LogFile $v_logPathfile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "Archive retrieval output folder : $v_logPathfile."


#We look at the input file line by line
ForEach ($affiliate in $affiliates) { 

    $piServerName = $affiliate.server

    # Connection to the PI server
    try 
    {
        Write-Log -v_LogFile $v_logPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$piServerName : Connecting to the current PI Server..."
        $con = Connect-PIDataArchive -PIDataArchiveMachineName $piServerName -AuthenticationMethod Windows -ErrorAction Stop
        Write-Log -v_LogFile $v_logPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$piServerName : Succesfull login !"
    }
    catch [System.Exception] 
    {
        # If connection KO, end of program
        Write-Log -v_LogFile $v_logPathfile -v_LogLevel WARN -v_ConsoleOutput -v_Message "Connection to the PI Server : $piServerName failed ...[$($_.Exception.Message)]"
        Write-Log -v_LogFile $v_logPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "End of application processing ..."
        Pause
        continue
    }

    Write-Host ""

    # Retrieving information from the archives.
    if($corruptOnly -eq "n")
    {
        Write-Log -v_LogFile $v_logPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Recovery of the entire archive in progress ..."
        (Get-PIArchiveInfo  -Connection $con).ArchiveFileInfo | Select-Object -Property Path,IsCorrupt | Out-File -FilePath $outputFile
        Write-Log -v_LogFile $v_logPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "Successful archive list retrieval !"
    }
    else 
    {
        Write-Log -v_LogFile $v_logPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Recovery of corrupted archives in progress ..."
        (Get-PIArchiveInfo  -Connection $con).ArchiveFileInfo | Select-Object -Property Path,IsCorrupt | Where-object {$_isCorrupt -eq $false} | Out-File -FilePath $outputFolder
        Write-Log -v_LogFile $v_logPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "Successful corrupted archive list retrieval !"
    }
 
}

Write-Host ""

Write-Log -v_LogFile $v_logPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Thank you for using PIArchiveList."

Pause