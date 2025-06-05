<#
.SYNOPSIS
Backfill the tags from a source server to a target server

.DESCRIPTION
Backfill the tags from a source server to a target server

.PARAMETER PISourceServer
[Mandatory] Name of the PI Server to extract data. Default is 'PISRV01'.

.PARAMETER PITargetServer
[Mandatory] Name of the PI Server to insert data. Default is 'PISRV02'.

.PARAMETER DateStartTime
[Mandatory] Start time of extraction (Format : yyyy-MM-ddThh:mm). Default is one month ago.

.PARAMETER DateEndTime
[Mandatory] End time of extraction (Format : yyyy-MM-ddThh:mm). Default is now.

.NOTES
Title: BackfillingPS
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Date v1 : 19/10/2022
Date v2 : 29/08/2024
Version: 2.1 : 05/06/2025
#>

Param(
    [Parameter(Mandatory=$False, HelpMessage="Name of the PI Server to extract data")]
    [String]$PISourceServer = 'PISRV01',

    [Parameter(Mandatory=$False, HelpMessage="Name of the PI Server to insert data")]
    [String]$PITargetServer = 'PISRV02',

    [Parameter(Mandatory=$False, HelpMessage="Start time of extraction (Format : yyyy-MM-ddThh:mm)")]
    [String]$DateStartTime = ((Get-Date).AddMonths(-1).ToString('yyyy-MM-ddTHH:mm')),

    [Parameter(Mandatory=$False, HelpMessage="End time of extraction (Format : yyyy-MM-ddThh:mm)")]
    [String]$DateEndTime = (Get-Date).ToString('yyyy-MM-ddTHH:mm')
)

# Import necessary modules and clear host
Clear-Host
Import-Module (Join-Path $PSScriptRoot '..\lib\logs.psm1') 
Import-Module (Join-Path $PSScriptRoot '..\lib\connection.psm1')
Import-Module (Join-Path $PSScriptRoot '..\lib\files.psm1')

# Function to initialize PI connections
function Initialize-PIConnections {
    param (
        [string]$PISourceServer,
        [string]$PITargetServer
    )
    $PIConSource = Connect-PIServer -PIServerHost $PISourceServer
    $PIConTarget = Connect-PIServer -PIServerHost $PITargetServer
    try {
        $PICollective = Get-PICollective -Connection $PIConTarget -ErrorAction Stop
        $PIConTargetPrimary = Connect-PIServer -PIServerHost $PICollective.Members.Name[0]
        $PIConTargetSecondary = Connect-PIServer -PIServerHost $PICollective.Members.Name[1]
        $IsCollective = $true
    } catch {
        $PIConTargetPrimary = $null
        $PIConTargetSecondary = $null
        $IsCollective = $false
    }
    return @{
        "Source" = $PIConSource;
        "Target" = $PIConTarget;
        "Primary" = $PIConTargetPrimary;
        "Secondary" = $PIConTargetSecondary;
        "IsCollective" = $IsCollective
    }
}

# Function to update a single tag
function Update-TagData {
    param (
        [string]$Tag,
        [DateTime]$DateStartTime,
        [DateTime]$DateEndTime,
        [int]$TagCounter,
        [hashtable]$PIConnections,
        [int]$TotalTags,
        [string]$LogPath,
        [double]$TotalTimeSpanInMinutes
    )

    Write-Log -LogLevel INFO -Message "$Tag ($TagCounter/$TotalTags) - Tag processing ..."
    $CurrentStartDate = $DateStartTime
    $Point = Get-PIpointSafe -TagName $Tag -PIConnection $PIConnections["Source"]

    while ($CurrentStartDate -lt $DateEndTime) {
        $RemainingTimeSpanInMinutes = (New-TimeSpan -Start $CurrentStartDate -End $DateEndTime).TotalMinutes
        Write-Progress -Activity "Processing backfilling for the period ($DateStartTime >> $DateEndTime)" -Status "Current period is $CurrentStartDate" -ParentId 1 -PercentComplete (($TotalTimeSpanInMinutes - $RemainingTimeSpanInMinutes) / $TotalTimeSpanInMinutes * 100)

        $CurrentEndDate = if ($CurrentStartDate.AddDays(1) -lt $DateEndTime) { 
            $CurrentStartDate.AddDays(1) 
        } else { 
            $DateEndTime 
        }

        $PIData = Get-PIValuesSafe -PIPoint $Point -StartTime $CurrentStartDate -EndTime $CurrentEndDate
        
        foreach ($CurrentData in $PIData) {
            if(($null -eq $CurrentData.Value) -or ($CurrentData.Value -eq "")) {$CurrentData.Value = " "}
            if ($PIConnections["IsCollective"]) {
                Add-PIValue -WriteMode "Replace" -Connection $PIConnections["Primary"] -PointName $Tag -Time $CurrentData.Timestamp -Value $CurrentData.Value -ErrorAction Stop | Out-Null
                Add-PIValue -WriteMode "Replace" -Connection $PIConnections["Secondary"] -PointName $Tag -Time $CurrentData.Timestamp -Value $CurrentData.Value -ErrorAction Stop | Out-Null
            } else {
                Add-PIValue -WriteMode "Replace" -Connection $PIConnections["Target"] -PointName $Tag -Time $CurrentData.Timestamp -Value $CurrentData.Value -ErrorAction Stop | Out-Null
            }
        }
        $CurrentStartDate = $CurrentEndDate
    }
    Write-Log -LogLevel INFO -Message "$Tag ($TagCounter/$TotalTags) - Tag processing ..."
}

# Main function to orchestrate the script
function Invoke-Backfill {
    $LogPath = Join-Path $PSScriptRoot "logs\logs_$((Get-Date).ToString('yyyy-MM-ddThh-mm-ss')).txt"
    $SourceFile = Join-Path $PSScriptRoot "input\points.txt"
    $Points = Get-Content $SourceFile
    $TotalTags = ($Points | Measure-Object -Line).Lines
    $DateStartTime = [datetime]::ParseExact($DateStartTime, "yyyy-MM-ddTHH:mm", $null)
    $DateEndTime = [datetime]::ParseExact($DateEndTime, "yyyy-MM-ddTHH:mm", $null)
    $TotalTimeSpanInMinutes = (New-TimeSpan -Start $DateStartTime -End $DateEndTime).TotalMinutes

    $PIConnections = Initialize-PIConnections -PISourceServer $PISourceServer -PITargetServer $PITargetServer
    $TagCounter = 0

    foreach ($Point in $Points) {
        $TagCounter++
        Write-Progress -Activity "Processing backfilling of $TotalTags tags from $PISourceServer to $PITargetServer. Is it a collective server: $($PIConnections["IsCollective"])" -Status "Processing tag $Point ($TagCounter/$TotalTags)" -Id 1 -PercentComplete ($TagCounter / $TotalTags * 100)
        Update-TagData -Tag $Point -DateStartTime $DateStartTime -DateEndTime $DateEndTime -TagCounter $TagCounter -PIConnections $PIConnections -TotalTags $TotalTags -LogPath $LogPath -TotalTimeSpanInMinutes $TotalTimeSpanInMinutes
    }

    Write-Log -LogLevel INFO -Message "End of Backfilling."
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("End of Backfilling.", "PI-DA_Backfilling", 0, 64)
}

# Call the main function
Invoke-Backfill