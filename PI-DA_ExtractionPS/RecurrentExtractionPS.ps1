<#
.SYNOPSIS
Specify to the original ExtractionPS script a date based on realtime
 
 
.DESCRIPTION
Specify to the original ExtractionPS script a date based on realtime
 
 
.PARAMETER PIServerHost
[Mandatory] Name of the PI Server to extract tags.
 
 
.PARAMETER startTime
[Mandatory] Start time of extraction.


.PARAMETER endTime
[Mandatory] End time of extraction.


.PARAMETER timeZone
[Optional] UTC option (localtime as default).


.PARAMETER doCompress
[Optional] Compression option of the output files.

.PARAMETER doCompressAll
[Optional] Compression option of the output files in one zip.


.PARAMETER noEmptyFile
[Optional] No empty files generated option.
 
.NOTES
Title: ExtractionPS
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Version: 0.1
#> 

param (
    [Parameter(Mandatory=$false, HelpMessage="Name of the PI Server to extract tags (exemple : PI-CENTER-HQ)")][int]$offset = 0, #0s in param section,
    [Parameter(Mandatory=$false, HelpMessage="Name of the PI Server to extract tags (exemple : PI-CENTER-HQ)")][int]$timeRange = 300, #5min in param section,
    [Parameter(Mandatory=$true, HelpMessage="Name of the PI Server to extract tags (exemple : PI-CENTER-HQ)")][String]$PIServerHost, #ExtractionPS parameter
    [Parameter(Mandatory=$true, HelpMessage="Output folder, default is output\ ")][String]$output, #ExtractionPS parameter
	[Parameter(Mandatory=$False, HelpMessage="Option to compress the files after extraction ?")][Switch]$doCompress,
	[Parameter(Mandatory=$False, HelpMessage="Option to compress the files after extraction in one zip?")][Switch]$doCompressAll,
    [Parameter(Mandatory=$False, HelpMessage="Option to generate empty files ?")][Switch]$noEmptyFile
)

#$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList @($username,(ConvertTo-SecureString -String $password -AsPlainText -Force))

#Script ExtractionPS location
$scriptPath = (Join-Path $PSScriptRoot "ExtractionPS.ps1")
#Now
$startDate = Get-Date
#write-output $startDate
# Now - offset
$startDate = $startDate.AddSeconds(-$offset)
# Now - offset - timeRange
$endDate = $startDate.AddSeconds(-$timeRange)

#Convert date to string
$startDate = $startDate.toString('yyyy-MM-ddThh:mm:ss')
$endDate = $endDate.toString('yyyy-MM-ddThh:mm:ss')

$argument = "-file .\ExtractionPS.ps1 -PIServerHost $PIServerHost -startTime $startDate -endTime $endDate -output $output"

if($doCompress){
	$argument = $argument + " -doCompress"
}

if($doCompressAll){
	$argument = $argument + " -doCompressAll"
}

if($noEmptyFile){
	$argument = $argument + " -noEmptyFile"
}

Start-Process powershell.exe -argumentList $argument





