<#
.SYNOPSIS
This script insert data contained into the input file to the choosen PI server.
 
 
.DESCRIPTION
This script insert data contained into the input file to the choosen PI server.
 
 
.PARAMETER piServerHost
[MANDATORY] PI DA Server where the tags will be provided in data.
  
 
.NOTES
Title: InsertionPS
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Version: 0.1
#> 


Param(
#[Parameter(Mandatory=$false, HelpMessage="Name of the PI Server to create tags (exemple : PI-CENTER-HQ)")][String]$piServerHost = "OPEPPA-WRPIHQ01"
[Parameter(Mandatory=$true, HelpMessage="Name of the PI Server to create tags (exemple : PI-CENTER-HQ)")][String]$piServerHost
)

import-module (Join-Path $PSScriptRoot '..\lib\Logs.psm1') 
import-module (Join-Path $PSScriptRoot '..\lib\connection.psm1')
import-module (Join-Path $PSScriptRoot '..\lib\regex.psm1')

#Initialization of variables and logs.
$timestamp=get-date -uFormat "%m%d%Y%H%M"
$v_LogPathfile = (Join-Path $PSScriptRoot "logs\logs_$timestamp.txt") 
$sourceFolder = (Join-Path $PSScriptRoot "input\")
Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

#Loading data files
$filesName = Get-ChildItem $sourceFolder -Filter Extraction*
$fileNumber = ($filesName | Measure-Object -line).Lines

#Connection to the PI server
$PIConnection = Connect-PIServer -PIServerHost $PIServerHost

#RITM2364818 - Write in all members of a collective server
try
{
    $PICollective = Get-PICollective $PIConnection -ErrorAction Stop
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "$PIServerHost is a collective server - InsertionPS will write on both members"
    $IsACollective = 1
}
catch
{
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "$PIServerHost is a single server - InsertionPS will write on this server only"
    $IsACollective = 0
}

if($IsACollective)
{
    Foreach ($PIServerMember in $PICollective.Members.Name)
    {
        #Connection to the PI server
        $PIConnection = Connect-PIServer -PIServerHost $PIServerMember
        
        $fileCounter = 0
        Clear-Host
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "Writing to member $PIServerMember."

        Foreach ($fileName in $filesName)
        {
            $fileCounter++
            Write-Progress -Activity ("PI-DA_InsertionPS - Treatment of $fileNumber files - Current PIServerMember is $PIServerMember") -Status "Processing file N°$fileCounter/$fileNumber ..." -id 2 -PercentComplete (($fileCounter-1)*100/$fileNumber)
            
            #Import CSV from current file
            try
            {
                Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Import-Csv file $fileName in progress ..."
                $currentCsvFile = Import-Csv -Path $sourceFolder$fileName -Delimiter ';' -Header tag,timestamp,value,quality -Encoding UTF8 -ErrorAction stop
                Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "File $fileName loaded."

                $nbline = (get-content $sourceFolder$fileName | Measure-Object -line).Lines
                #$currentCsvFile.length
                $currentline=0
            }
            catch [System.Exception]
            {
                Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$fileName : Error during Import-Csv function to this file. [$($_.Exception.Message)]"
            }
        
            Foreach($line in $currentCsvFile)
            {
                $currentline++
                Write-Progress -Activity ("PI-DA_InsertionPS - Insertion of data in $fileName") -Status "Processing" -id 3 -ParentId 2 -PercentComplete ($currentline*100/$nbline)
                #Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_Message "Processing : $($line.tag)- $($line.timestamp)- $($line.value)- $($line.quality)"
                try
                {
                    #RITM2364818 - Distinct insertion for Good & No Good values to insert "good" digital values without issues
                    if($line.quality -eq "Good Value")
                    {
                        if($line.value.startswith("State"))
                        {
                            #Get the length of the numerical value of the digital state
                            $NumericLength = Get-NumericalLength($line.value)
                            Add-PIValue -WriteMode "Replace" -Connection $PIConnection -PointName $line.tag -Time $line.timestamp -Value $line.value.substring(7,$NumericLength) -ErrorAction Stop | Out-Null
                        }
                        else
                        {
                            Add-PIValue -WriteMode "Replace" -Connection $PIConnection -PointName $line.tag -Time $line.timestamp -Value $line.value -ErrorAction Stop | Out-Null
                        }
                    }
                    elseif($line.quality -eq "No Good Value")
                    {
                        Add-PIValue -WriteMode "Replace" -Connection $PIConnection -PointName $line.tag -Time $line.timestamp -Value $line.value.substring(7,3) -UseSystemState 1 -ErrorAction Stop | Out-Null
                    }
                    else #Input file not good
                    {
                        throw "Quality of the tag not specified for this line - Please use latest version of ExtractionPS to get the good input files."
                    }
                }
                catch [System.Exception]
                {
                    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "File : $fileName. Error providing data to the tag. [$($_.Exception.Message)]"
                }
            } 
            Remove-Variable currentCsvFile
        }
        Disconnect-PIDataArchive -Connection $PIConnection | Out-Null
        while($PIConnection.Connected) {Start-Sleep -Milliseconds 500}
    }
}
else #not a collective server
{
    Foreach ($fileName in $filesName)
    {
        $fileCounter++
        Write-Progress -Activity ("PI-DA_InsertionPS - Treatment of $fileNumber files") -Status "Processing file N°$fileCounter ..." -id 2 -PercentComplete ($fileCounter*100/$fileNumber)
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "Processing file $fileName ($fileCounter/$fileNumber)"
    
        #Import CSV from current file
        try
        {
            $currentCsvFile = Import-Csv -Path $sourceFolder$fileName -Delimiter ';' -Header tag,timestamp,value,quality -Encoding UTF8 -ErrorAction stop
            #$nbline = ($currentCsvFile | Measure-Object -line).Lines
            $nbline = (get-content $sourceFolder$fileName | Measure-Object -line).Lines
            $currentline=0
        }
        catch [System.Exception]
        {
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$fileName : Error during Import-Csv function to this file. [$($_.Exception.Message)]"
        }
    
        Foreach($line in $currentCsvFile)
        {
            $currentline++
            Write-Progress -Activity ("PI-DA_InsertionPS - Insertion of data in $fileName") -Status "Processing" -id 3 -ParentId 2 -PercentComplete ($currentline*100/$nbline)
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_Message "Processing : $($line.tag)- $($line.timestamp)- $($line.value)- $($line.quality)"
            try
            {
                #RITM2364818 - Distinct insertion for Good & No Good values to insert "good" digital values without issues
                if($line.quality -eq "Good Value")
                {
                    if($line.value -like "State*")
                    {
                        #Get the length of the numerical value of the digital state
                        $NumericLength = Get-NumericalLength($line.value)
                        Add-PIValue -WriteMode "Replace" -Connection $PIConnection -PointName $line.tag -Time $line.timestamp -Value $line.value.substring(7,$NumericLength) -ErrorAction Stop  
                    }
                    else{
                        Add-PIValue -WriteMode "Replace" -Connection $PIConnection -PointName $line.tag -Time $line.timestamp -Value $line.value -ErrorAction Stop  
                    }
                }
                elseif($line.quality -eq "No Good Value")
                {
                    Add-PIValue -WriteMode "Replace" -Connection $PIConnection -PointName $line.tag -Time $line.timestamp -Value $line.value.substring(7,3) -UseSystemState 1 -ErrorAction Stop
                }
                else #Input file not good
                {
                    throw "Quality of the tag not specified for this line - Please use latest version of ExtractionPS to get the good input files."
                }
            }
            catch [System.Exception]
            {
                Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "File : $fileName. Error providing data to the tag. [$($_.Exception.Message)]"
            }
        } 
    }
    Disconnect-PIDataArchive -Connection $PIConnection | Out-Null
    while($PIConnection.Connected) {Start-Sleep -Milliseconds 500}
}


Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "End of insertion."
[System.Windows.Forms.MessageBox]::Show("End of insertion.", "PI-DA_InsertionPS", 0, 64)