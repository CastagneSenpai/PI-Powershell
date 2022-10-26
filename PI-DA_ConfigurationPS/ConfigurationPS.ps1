<#
.SYNOPSIS
Get the full configuration of tags in a PI Server.
 
 
.DESCRIPTION
Get the full configuration of tags in a PI Server.
 
 
.PARAMETER PIServerHost
[Mandatory] Name of the PI Server to extract tags.

 
.NOTES
Title: ConfigurationPS
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Version: 0.1
#> 


Param(
[Parameter(Mandatory=$True, HelpMessage="Name of the PI Server to extract tags (exemple : PI-CENTER-HQ)")][String]$PIServerHost = "AOEPTTA-APPIL01"
)

import-module (Join-Path $PSScriptRoot '..\lib\Logs.psm1') 


#---------------------------------------------------------------------------------#
#Initialization of variables, logs.
$timestamp=get-date -uFormat "%Y%m%d_%Hh%M"
$fileDataPath = Join-Path $PSScriptRoot "output\tagConfiguration_$timestamp.csv"
$v_LogPathfile = Join-Path $PSScriptRoot "logs\logs_$timestamp.txt"
$sourceFile = Join-Path $PSScriptRoot "input\points.txt"
Clear-Host
#---------------------------------------------------------------------------------#


#---------------------------------------------------------------------------------#
#Connexion to the PI server
try 
{
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "$PIServerHost : Connexion to the PI server in progress ..."
    $piConnexion = Connect-PIDataArchive -PIDataArchiveMachineName $PIServerHost -AuthenticationMethod Windows -ErrorAction Stop
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$PIServerHost : Connexion to server successful."
}
catch [System.Exception] 
{
    #If KO Connexion, end of program
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$PIServerHost : Connexion to server failed. [$($_.Exception.Message)]"
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "End of application processing."
    Pause
    exit
}
#---------------------------------------------------------------------------------#

#Loading tags in the input file
$points = Get-Content $sourceFile 

#Tag counter
$iTagCourant = 0
$nbTagsTotal =  ($points | measure-object -line).Lines

#Creation of the file stream to write on the output file  
$fs = New-Object System.IO.FileStream $fileDataPath ,'Append','Write','Read' 
$myStreamWriter =  New-Object System.IO.StreamWriter($fs)

#Header of file
$headerAttributes = "convers;compmin;zero;compmax;ptclassid;compdevpercent;userint1;excmin;dataaccess;squareroot;step;scan;location2;location3;location1;excmax;changer;ptgroup;ptclassrev;tag;excdevpercent;ptaccess;recno;creator;descriptor;digitalset;filtercode;pointsource;archiving;typicalvalue;span;ptowner;excdev;ptsecurity;datagroup;location5;userreal1;displaydigits;pointtype;instrumenttag;creationdate;userint2;changedate;exdesc;userreal2;shutdown;datasecurity;compressing;engunits;ptclassname;pointid;location4;srcptid;compdev;dataowner;totalcode;sourcetag"
$myStreamWriter.WriteLine($headerAttributes)

#---------------------------------------------------------------------------------#
foreach($point in $points)
{
    #Get current tag configuration
    try
    {
        $config = Get-PIpoint -Name $point -AllAttributes -Connection $piConnexion -ErrorAction Stop       
    }
    catch [System.Exception]
    {
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$point : Error using Get-PIpoint function for this tag. [$($_.Exception.Message)]"
        Read-Host -Prompt "Press Enter to quit ConfigurationPS."
        exit
    }
    
    #Write into output file current tag configuration
    try
    {
        $attributesList = $config.Attributes
        $myStreamWriter.WriteLine(
            "$($attributesList.convers);" + 
            "$($attributesList.compmin);" + 
            "$($attributesList.zero);" + 
            "$($attributesList.compmax);" + 
            "$($attributesList.ptclassid);" + 
            "$($attributesList.compdevpercent);" + 
            "$($attributesList.userint1);" + 
            "$($attributesList.excmin);" + 
            "$($attributesList.dataaccess);" + 
            "$($attributesList.squareroot);" + 
            "$($attributesList.step);" + 
            "$($attributesList.scan);" + 
            "$($attributesList.location2);" + 
            "$($attributesList.location3);" + 
            "$($attributesList.location1);" + 
            "$($attributesList.excmax);" + 
            "$($attributesList.changer);" + 
            "$($attributesList.ptgroup);" + 
            "$($attributesList.ptclassrev);" + 
            "$($attributesList.tag);" + 
            "$($attributesList.excdevpercent);" + 
            "$($attributesList.ptaccess);" + 
            "$($attributesList.recno);" + 
            "$($attributesList.creator);" + 
            "$($attributesList.descriptor);" + 
            "$($attributesList.digitalset);" + 
            "$($attributesList.filtercode);" + 
            "$($attributesList.pointsource);" + 
            "$($attributesList.archiving);" + 
            "$($attributesList.typicalvalue);" + 
            "$($attributesList.span);" + 
            "$($attributesList.ptowner);" + 
            "$($attributesList.excdev);" + 
            "$($attributesList.ptsecurity);" + 
            "$($attributesList.datagroup);" + 
            "$($attributesList.location5);" + 
            "$($attributesList.userreal1);" + 
            "$($attributesList.displaydigits);" + 
            "$($attributesList.pointtype);" + 
            "$($attributesList.instrumenttag);" + 
            "$($attributesList.creationdate);" + 
            "$($attributesList.userint2);" + 
            "$($attributesList.changedate);" + 
            "$($attributesList.exdesc);" + 
            "$($attributesList.userreal2);" + 
            "$($attributesList.shutdown);" + 
            "$($attributesList.datasecurity);" + 
            "$($attributesList.compressing);" + 
            "$($attributesList.engunits);" + 
            "$($attributesList.ptclassname);" + 
            "$($attributesList.pointid);" + 
            "$($attributesList.location4);" + 
            "$($attributesList.srcptid);" + 
            "$($attributesList.compdev);" + 
            "$($attributesList.dataowner);" + 
            "$($attributesList.totalcode);" + 
            "$($attributesList.sourcetag)"
         )
    }
    catch [System.Exception]
    {
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$point : Error writing result of Get-PIpoint function to $fileDataPath. [$($_.Exception.Message)]"
        Read-Host -Prompt "Press Enter to quit ConfigurationPS."
        exit
    }
}
#---------------------------------------------------------------------------------#

$myStreamWriter.Close()
Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "Tags configuration available at $fileDataPath."
Read-Host -Prompt "Press Enter to quit ConfigurationPS."