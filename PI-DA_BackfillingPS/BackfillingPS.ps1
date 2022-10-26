<#
.SYNOPSIS
Backfill the tags from a source server to a target server

.DESCRIPTION
Backfill the tags from a source server to a target server
 
.PARAMETER PISourceServer
[Mandatory] Name of the PI Server to extract data.
 
.PARAMETER PITargetServer
[Mandatory] Name of the PI Server to insert data.

.PARAMETER DateStartTime
[Mandatory] Start time of backfilling.

.PARAMETER DateEndTime
[Mandatory] End time of backfilling.

.NOTES
Title: BackfillingPS
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Date : 19/10/2022
Version: 0.1
#> 

Param(
[Parameter(Mandatory=$True, HelpMessage="Name of the PI Server to extract data")][String]$PISourceServer,
[Parameter(Mandatory=$True, HelpMessage="Name of the PI Server to insert data")][String]$PITargetServer,
[Parameter(Mandatory=$True, HelpMessage="Start time of extraction (Format : yyyy-MM-ddThh:mm)")][String]$DateStartTime,
[Parameter(Mandatory=$True, HelpMessage="End time of extraction (Format : yyyy-MM-ddThh:mm)")][String]$DateEndTime
)

Clear-Host
import-module (Join-Path $PSScriptRoot '..\lib\logs.psm1') 
import-module (Join-Path $PSScriptRoot '..\lib\connection.psm1')
import-module (Join-Path $PSScriptRoot '..\lib\files.psm1')

#Init the logfile
$v_LogPathfile = Join-Path $PSScriptRoot "logs\logs_$((get-date).toString('yyyy-MM-ddThh-mm-sstt')).txt"

#Convert dates in input from [string] to [Date]
 [DateTime]$DateStartTime = [datetime]::ParseExact($DateStartTime, "yyyy-MM-ddThh:mm", $null)
 [DateTime]$DateEndTime = [datetime]::ParseExact($DateEndTime, "yyyy-MM-ddThh:mm", $null)

#Loading tags in the input file
$sourceFile = Join-Path $PSScriptRoot "input\points.txt"
$points = Get-Content $sourceFile

#Setup variables for the progress bar
$TagCounter = 0 #Tag counter
$NumbersOfTags =  ($points | measure-object -line).Lines
$TotalTimeSpanInMinutes = (New-TimeSpan -start $DateStartTime -end $DateEndTime).TotalMinutes

#Connection to the PI source & PI target server - check if target server is a collective to replicate on both
$PIConSource = Connect-PIServer -PIServerHost $PISourceServer
$PIConTarget = Connect-PIServer -PIServerHost $PITargetServer
try {
    $PICollective = Get-PICollective $PIConTarget -ErrorAction Stop
    $PIConTargetPrimary = Connect-PIServer -PIServerHost $PICollective.Members.Name[0]
    $PIConTargetSecondary = Connect-PIServer -PIServerHost $PICollective.Members.Name[1]
    $IsACollective = $true
} catch {
    $IsACollective = $false
}

#TAG EXTRACTION
ForEach ($point in $points)
{
    #Update progress bar 1
    Write-Progress -Activity ("Processing backfilling of $NumbersOfTags tags from $PISourceServer to $PITargetServer. Is it a collective server : $IsACollective") -Status "Processing tag $point ($TagCounter/$NumbersOfTags)" -id 1 -PercentComplete ($TagCounter/$NumbersOfTags*100)
    $TagCounter++ #Increment of the tag counter
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_message "$point ($TagCounter)/$NumbersOfTags) - Tag processing ..."

    #Set a mobile date variable
    $CurrentStartDate = $DateStartTime
    
    #Get the current PIpoint to backfill
    $pt = Get-PIpointSafe -TagName $point -PIConnection $PIConSource
    Select-Object -Property timestamp, Value

    #Backfilling data from the current tag
    while($CurrentStartDate -lt $DateEndTime) # -lt : inférieur à
    {
        #Update progress bar 2
        $RemainingTimeSpanInMinutes = (New-TimeSpan -start $CurrentStartDate -end $DateEndTime).TotalMinutes
        Write-Progress -Activity ("Processing backfilling for the period ($DateStartTime >> $DateEndTime)") -Status "Current period is $CurrentStartDate" -ParentId 1 -PercentComplete (($TotalTimeSpanInMinutes-$RemainingTimeSpanInMinutes)/$TotalTimeSpanInMinutes*100)

        #Calculate the current end date of this micro extraction : CurrentStartDate + One day 
        if($CurrentStartDate.AddDays(1) -lt $DateEndTime){
            $CurrentEndDate = $CurrentStartDate.AddDays(1)
        } else {
            $CurrentEndDate = $DateEndTime
        }
        
        #Get data for a day
        $PIData = Get-PIValuesSafe -PIPoint $pt -st $CurrentStartDate -et $CurrentEndDate
        
        #Insert PIData in the target server tag
        foreach($CurrentData in $PIData)
        {
            if($IsACollective){
                Add-PIValue -WriteMode "Replace" -Connection $PIConTargetPrimary -PointName $point -Time $CurrentData.Timestamp -Value $CurrentData.Value -ErrorAction Stop | Out-Null
                Add-PIValue -WriteMode "Replace" -Connection $PIConTargetSecondary -PointName $point -Time $CurrentData.Timestamp -Value $CurrentData.Value -ErrorAction Stop | Out-Null
            } else {
                Add-PIValue -WriteMode "Replace" -Connection $PIConTarget -PointName $point -Time $CurrentData.Timestamp -Value $CurrentData.Value -ErrorAction Stop | Out-Null
            }   
        }
        
        #Update the variable for next micro extraction
        $CurrentStartDate = $CurrentEndDate

        #Update progress bar 2
        $RemainingTimeSpanInMinutes = (New-TimeSpan -start $CurrentStartDate -end $DateEndTime).TotalMinutes
        Write-Progress -Activity ("Processing backfilling for the period ($DateStartTime >> $DateEndTime)") -Status "Current period is $CurrentStartDate" -ParentId 1 -PercentComplete (($TotalTimeSpanInMinutes-$RemainingTimeSpanInMinutes)/$TotalTimeSpanInMinutes*100)
    }
    #Update progress bar 1
    Write-Progress -Activity ("Processing backfilling of $NumbersOfTags tags from $PISourceServer to $PITargetServer") -Status "Processing tag $point ($TagCounter/$NumbersOfTags)" -id 1 -PercentComplete ($TagCounter/$NumbersOfTags*100)
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_message "$point ($TagCounter/$NumbersOfTags) - Tag processing ..."
}

Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "End of Backfilling."
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("End of Backfilling.", "PI-DA_Backfilling", 0, 64)