<#
.SYNOPSIS
This script create tags structure listed in the input file to the choosen PI server.
 
 
.DESCRIPTION
This script create tags structure listed in the input file to the choosen PI server.
 
 
.PARAMETER piServerHost
[MANDATORY] PI DA Server where the tags will be created.
  
 
.NOTES
Title: CreationPS
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Version: 0.1
#> 


Param(
#[Parameter(Mandatory=$false, HelpMessage="Name of the PI Server to create tags (exemple : PI-CENTER-HQ)")][String]$piServerHost = "OPEPPA-WRPIHQ01"
[Parameter(Mandatory=$true, HelpMessage="Name of the PI Server to create tags (exemple : PI-CENTER-HQ)")][String]$piServerHost
)

import-module (Join-Path $PSScriptRoot '..\lib\Logs.psm1') 

#Initialization of variables and logs.
$timestamp=get-date -uFormat "%m%d%Y%H%M"
$v_LogPathfile = (Join-Path $PSScriptRoot "logs\logs_$timestamp.txt") 
$sourceFolder = (Join-Path $PSScriptRoot "input\")
Clear-Host

#Loading ConfigurationPS files
$filesName = (Join-Path $sourceFolder (Get-ChildItem $sourceFolder -Filter tagConfiguration_*))

#Connection to the PI server
try 
{
    $piConnection = Connect-PIDataArchive -PIDataArchiveMachineName $piServerHost -AuthenticationMethod Windows -ErrorAction Stop
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$piServerHost : Connected to the server."
}
catch [System.Exception] 
{
    #If KO connection, end of program
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$piServerHost : Connection to server failed. [$($_.Exception.Message)]."
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "End of application processing."
    Pause
    exit
}

Foreach ($fileName in $filesName)
{
    $currentLine = 0
    #Import CSV from current file
    try
    {
        $currentCsvFile = Import-Csv -Path $fileName -Delimiter ';' -Encoding UTF8 -ErrorAction stop
    }
    catch [System.Exception]
    {
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$fileName : Error during Import-Csv function to this file. [$($_.Exception.Message)]"
    }

    Foreach($line in $currentCsvFile)
    {
        $currentLine++
        try
        {
            Add-PIPoint -Connection $piConnection -Name $line.tag -Attributes @{
            archiving="$($line.archiving)";
            compdev="$($line.compdev)";
            compdevpercent="$($line.compdevpercent)";
            compmax="$($line.compmax)";
            compmin="$($line.compmin)";
            compressing="$($line.compressing)";
            convers="$($line.convers)";
            #dataaccess="$($line.dataaccess)";
            datasecurity="$($line.datasecurity)";
            descriptor="$($line.descriptor)";
            digitalset="$($line.digitalset)";
            displaydigits="$($line.displaydigits)";
            engunits="$($line.engunits)";
            excdev="$($line.excdev)";
            excdevpercent="$($line.excdevpercent)";
            excmax="$($line.excmax)";
            excmin="$($line.excmin)";
            exdesc="$($line.exdesc)";
            filtercode="$($line.filtercode)";
            instrumenttag="$($line.instrumenttag)";
            location1="$($line.location1)";
            location2="$($line.location2)";
            location3="$($line.location3)";
            location4="$($line.location4)";
            location5="$($line.location5)";
            pointsource="$($line.pointsource)";
            pointtype="$($line.pointtype)";
            #ptaccess="$($line.ptaccess)";
            ptclassid="$($line.ptclassid)";
            ptclassname="$($line.ptclassname)";
            ptsecurity="$($line.ptsecurity)";
            scan="$($line.scan)";
            shutdown="$($line.shutdown)";
            sourcetag="$($line.sourcetag)";
            span="$($line.span)";
            squareroot="$($line.squareroot)";
            srcptid="$($line.srcptid)";
            step="$($line.step)";
            typicalvalue="$($line.typicalvalue)";
            userint1="$($line.userint1)";
            userint2="$($line.userint2)";
            userreal1="$($line.userreal1)";
            userreal2="$($line.userreal2)";
            } -ErrorAction stop | Out-Null
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Line ($currentLine) : $($line.tag) created"
        }

        catch [System.Exception]
        {
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "File : $fileName. Line ($currentLine) : Error creating tag because some attributes are not correct or tag already exist.  [$($_.Exception.Message)]"
        }
    } 
}

Disconnect-PIDataArchive -Connection $piConnection | Out-Null
Read-Host -Prompt "Press Enter to quit CreationPS."