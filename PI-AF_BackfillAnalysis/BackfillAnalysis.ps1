<#
.SYNOPSIS
    Script to manage AFAnalysis backfill processes and PI AF analysis operations.

.DESCRIPTION
    This script automates the backfilling of AFAnalysis in PI AF using OSIsoft AF SDK. It includes connection management to AF servers, CSV-based input for analyses, log handling, and configurable time ranges for the backfill. 
    It also handles the start and stop of the specified analyses.

.NOTES
    Developer  : Romain Castagné
    Email      : romain.castagne-ext@syensqo.com
    Company    : CGI - SYENSQO
    Date       : 23/10/2024
    Version    : 1.0

.PARAMETER afServerName
    The name of the PI AF server to connect to.

.PARAMETER afDBName
    The name of the PI AF database to use.

.PARAMETER afSDKPath
    Path to the OSIsoft AF SDK DLL.

.PARAMETER InputCsvFilePathAndName
    Path and filename for the input CSV file containing analysis details.

.PARAMETER DeltaStartInMinutes
    Time range start offset in minutes from the current time (for backfill).

.PARAMETER DeltaEndInMinutes
    Time range end offset in minutes from the current time (for backfill).

.PARAMETER AutomaticMode
    Set to true for automatic pauses, false for manual validation.

#>


param(
        [string]$afServerName = "vmcegdidev001",
        [string]$afDBName = "Romain_Dev",
        [string]$afSDKPath = "E:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0\OSIsoft.AFSDK.dll",
        [string]$RecalculationLogFilePath = "C:\ProgramData\OSIsoft\PIAnalysisNotifications\Data\Recalculation\recalculation-log.csv",
        [string]$InputCsvFilePathAndName = (Join-path $PSScriptRoot "input.csv"),
        [int]$DeltaStartInMinutes = 1440,
        [int]$DeltaEndInMinutes = 30,
        [bool]$AutomaticMode = $true
)

# Log management function
function Write-Log {
    [CmdletBinding()]
    Param(
        [string]$v_Message,
        [string]$v_LogFile = (Join-Path -Path $PSScriptRoot -ChildPath ((Get-Date -Format yyyy-MM-dd) + "_Logs.txt")),
        [switch]$v_ConsoleOutput,
        [ValidateSet("SUCCESS", "INFO", "WARN", "ERROR", "DEBUG")]
        [string]$v_LogLevel = "INFO"
    )

    Begin {
        # Define log levels color
        $colorMap = @{
            "SUCCESS" = "Green"
            "INFO"    = "White"
            "WARN"    = "Yellow"
            "ERROR"   = "Red"
            "DEBUG"   = "Gray"
        }
        $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $color = $colorMap[$v_LogLevel]
    }

    Process {
        if ($v_Message) {
            $logEntry = "[$timeStamp] [$v_LogLevel] :: $v_Message"

            try {
                if ($v_LogFile) {
                    Out-File -Append -FilePath $v_LogFile -InputObject $logEntry
                }

                if ($v_ConsoleOutput.IsPresent) {
                    # [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                    Write-Host $logEntry -ForegroundColor $color
                }
            }
            catch {
                Write-Error "Failed to write log: $_"
            }
        }
    }
}


# Function to add the AFSDK library if not added previously
function Import-AFSDK{
    param(
        [string] $AFSDKPath = "C:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0\OSIsoft.AFSDK.dll"
    )
    if (-not (Test-Path $afSDKPath)) {
        Write-Log -v_Message "AF SDK not found at path: $afSDKPath" -v_ConsoleOutput -v_LogLevel ERROR
        exit
    }
    Add-Type -Path $afSDKPath
    Write-Log -v_Message "AF SDK found and loaded." -v_ConsoleOutput -v_LogLevel INFO
}

# Connection function to PI AF, returns the connection object of type [OSIsoft.AF.AFDatabase]
function Connect-AFServer {
    param (
        [string]$afServerName = "vmcegdidev001",
        [string]$afDBName = "Romain_Dev"
    )

    try {
        $afSystems = New-Object OSIsoft.AF.PISystems
        $afServer = $afSystems[$afServerName]
        $afServer.Connect()
        Write-Log -v_Message "Successfully connected to AF Server: $afServerName" -v_ConsoleOutput -v_LogLevel INFO
        $afDB = $afServer.Databases[$afDBName]

        if ($null -eq $afDB) {
            Write-Log -v_Message "Database $afDBName not found on AF Server $afServerName." -v_LogLevel ERROR -v_ConsoleOutput
        } else {
            Write-Log -v_Message "Successfully connected to AF Database: $afDBName" -v_ConsoleOutput -v_LogLevel INFO
        }
        return $afDB
    }
    catch {
        Write-Log -v_Message "Failed to connect to AF Server or Database: $_" -v_ConsoleOutput -v_LogLevel ERROR
        exit
    }
}

# Function for a clean disconnection from PI AF
function Disconnect-AFServer {
    param (
        [string]$afServerName = "vmcegdidev001"
    )

    try {
        $afSystems = New-Object OSIsoft.AF.PISystems
        $afServer = $afSystems[$afServerName]
        $afServer.Disconnect()
        Write-Log -v_Message "Successfully disconnected from AF Server: $afServerName" -v_ConsoleOutput -v_LogLevel INFO
    }
    catch {
        Write-log -v_Message "Failed to disconnect to AF Server: $_" -v_ConsoleOutput -v_LogLevel ERROR
        exit
    }
}

# Function that returns a list of AFAnalysis objects from a CSV file
function Get-AFAnalysisListFromCsvFile {
    param(
        $AFDatabase,
        [string] $InputCsvFilePathAndName = (Join-path $PSScriptRoot "Input.csv")
    )

    # Read the input file
    $FileContent = Import-Csv -Path $InputCsvFilePathAndName -Delimiter ','
    
    # Instanciate a list of AF Analysis 
    $AnalysisList = [System.Collections.Generic.List[OSIsoft.AF.Analysis.AFAnalysis]]::new()

    foreach($row in $fileContent){
        $AFAnalysisToAdd = $AFDatabase.Elements[$row.ElementPath].Analyses | Where-Object { $_.Name -eq $row.AnalysisName }
        if ($null -eq $AFAnalysisToAdd){
            Write-Log -v_Message "Analysis $(join-path $row.ElementPath $row.AnalysisName) not found in $($AFDatabase.Name)" -v_ConsoleOutput -v_LogLevel WARN
        }
        else{
            $AnalysisList.Add($AFAnalysisToAdd)
            Write-Log -v_ConsoleOutput -v_Message "--- Analysis `'$($row.ElementPath)\$AFAnalysisToAdd`' added for backfilling process."
        }
    }
    return $AnalysisList
}

# Function for constructing the AFTimeRange object from an offset and a desired backfill duration
function Format-AFTimeRange{
    param(
        [Int] $DeltaStartInMinutes = 1440, # A day per default
        [Int] $DeltaEndInMinutes = 30      # 30 minutes offset
    )

    $endTime = New-Object OSIsoft.AF.Time.AFTime((Get-Date).AddMinutes((-$DeltaEndInMinutes)))
    $startTime = New-Object OSIsoft.AF.Time.AFTime((Get-Date).AddMinutes(-$DeltaStartInMinutes))
    $afTimeRange = New-Object OSIsoft.AF.Time.AFTimeRange($startTime, $endTime)
    
    Write-Log -v_Message "The backfill time range is set to start at $($startTime) and end at $($endTime)." -v_ConsoleOutput -v_LogLevel INFO
    return $afTimeRange
}

# Function that combines a progress bar in VB with a Start-Sleep of X seconds
function Start-VisualSleep {
    param(
        [Int] $Seconds = 30,
        [String] $Activity = "Attente en cours"
    )

    # New form window
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Activity
    $form.Size = New-Object System.Drawing.Size(300, 100)
    
    # Create the progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 20)
    $progressBar.Size = New-Object System.Drawing.Size(250, 30)
    $progressBar.Style = 'Continuous'
    
    # Add progress bar to from
    $form.Controls.Add($progressBar)
    
    # Show the window
    $form.Show()

    # Define interval in milliseconds to make the progress bar fuild
    $interval = 100  
    $totalIterations = ($Seconds * 1000) / $interval 

    # Update the progress bar
    for ($i = 1; $i -le $totalIterations; $i++) {
        $progressBar.Value = ($i / $totalIterations) * 100  
        Start-Sleep -Milliseconds $interval 
    }

    # Close the window
    $form.Close()
}

# Function which allows you to wait until the backfill of an analysis list launched on a known timerange is completed (reading the recalculation-log.csv)
function Wait-EndOfBackfilling{
    param(
        [System.Collections.Generic.List[OSIsoft.AF.Analysis.AFAnalysis]] $AFAnalysisList,
        [OSIsoft.AF.Time.AFTimeRange] $AFTimeRange,
        [String] $RecalculationLogFilePath
    )

    # Initialize the analysis list with the status "InProgress"
    $AnalysisWithStatusList = @()
    foreach ($AFAnalysis in $AFAnalysisList) {
        $AnalysisWithStatusList += [PSCustomObject]@{
            Analysis = $AFAnalysis
            Status   = "InProgress"
        }
    }
    $AFTimeRangeConvertedToMatchLogFile = $AFTimeRange.StartTime.ToString("yyyy-MM-ddThh:00:00") + "--" + $AFTimeRange.EndTime.ToString("yyyy-MM-ddThh:00:00")
 
    # Loop to wait for backfilling to finish
    do {
        $RecalculationLogFile = Import-Csv -Delimiter ',' -Path $RecalculationLogFilePath
    
        foreach ($AnalysisWithStatus in $AnalysisWithStatusList) {
            if ($AnalysisWithStatus.Status -eq "Completed") {
                continue
            }
    
            # Check if a line in the log indicates that the analysis is complete or has an error
            foreach ($logLine in $RecalculationLogFile) {
    
                # Format the dates of the log file to compare with the sent AFTimeRange
                $LogDateArray = $logLine.TimeRange -split '--'
                $LogDateDebut = (Get-Date $LogDateArray[0]).ToString("yyyy-MM-ddThh:00:00")
                $LogDateFin = (Get-Date $LogDateArray[1]).ToString("yyyy-MM-ddThh:00:00")
                $LogLineTimeRangeFormatted = $LogDateDebut + "--" + $LogDateFin
    
                # Check that the log corresponds to the analysis and the specified period
                if ($logLine.Id -eq $AnalysisWithStatus.Analysis.UniqueID -and $LogLineTimeRangeFormatted -eq $AFTimeRangeConvertedToMatchLogFile -and $logLine.Type -eq "Manual") {
    
                    # Case when status is Completed
                    if ($logLine.Status -in ("Completed", "PendingCompletion")) {
                        Write-Log -v_Message "--- Backfilling of analysis $($AnalysisWithStatus.analysis.target)[$($AnalysisWithStatus.analysis.name)] completed." -v_LogLevel INFO -v_ConsoleOutput
                        $AnalysisWithStatus.Status = "Completed"
                        break
                    }
                    # Case when status is Error
                    elseif ($logLine.Status -eq "Error") {
                        Write-Log -v_Message "--- Error encountered in backfilling of analysis $($AnalysisWithStatus.analysis.target)[$($AnalysisWithStatus.analysis.name)] - Exiting." -v_LogLevel ERROR -v_ConsoleOutput
                        $AnalysisWithStatus.Status = "Error"
                        break
                    }
                }
            }    
            # Exit the main loop if an error was detected
            if ($AnalysisWithStatusList.Status -contains "Error") {
                return $false
            }
        }
    
        # Pause before next check
        Start-Sleep -Seconds 5
    
    } while ($AnalysisWithStatusList.Status -contains "InProgress" -and -not ($AnalysisWithStatusList.Status -contains "Error"))
}

# Function for launching a backfill of several analyzes on an AFTimeRange
function Start-AFAnalysisRecalculation{
    param(
        $AFDatabase,
        [System.Collections.Generic.List[OSIsoft.AF.Analysis.AFAnalysis]] $AFAnalysisList,
        [OSIsoft.AF.Time.AFTimeRange] $AFTimeRange,
        $AutomaticMode,
        [string] $RecalculationLogFilePath
    )

    $afAnalysisService = $AFDatabase.PISystem.AnalysisService
    try {  
        if($null -eq $AFAnalysisList){
            throw "There is no valid AFAnalysis in the input CSV file."
        }

        ## DEBUG : affiche le statut Disabled des analyses avant lancement
        # $AFAnalysisList | ForEach-Object { 
        #     $status = $_.GetStatus()
        #     Write-Host "Analyse: $($_.Name), Statut: $status"
        # }

        # Start AF Analysis
        Write-Log -v_Message "Starting all the analysis listed..." -v_ConsoleOutput -v_LogLevel INFO
        [OSIsoft.AF.Analysis.AFAnalysis]::SetStatus($AFAnalysisList, [OSIsoft.AF.Analysis.AFStatus]::Enabled)

        # Wait that analysis are well started
        if($AutomaticMode -eq $false ) { Read-Host "Pause - Validate Analysis started." }
        if($AutomaticMode -eq $true) { 

            # METHOD 1 - KO return empty : $AnalysisStatusList = [OSIsoft.AF.Analysis.AFAnalysis]::GetStatus($AFAnalysisList)
            # METHOD 2 - OK but return Enabled/Disabled instead of advanced fields (Starting, Running, Stopping, Stopped...)
            do {
                # Check that all analysis are Enabled
                $allEnabled = ($AFAnalysisList | ForEach-Object { $_.GetStatus() } | Where-Object { $_ -ne "Enabled" }).Count -eq 0
                
                # Retry 1 second later if not the case
                Start-Sleep -Seconds 1

            } while (-not $allEnabled)

            # METHOD 3 - Fixed pause : Start-VisualSleep -Seconds 12 -Activity "Analysis starting..." 
            
            # METHOD 4 : QueryRuntimeInformation -- Have to apply a filter to will be different from input csv analysis list + 'starting' status issue (cf. AVEVA Case)
            # Do{
            #     $results = $afAnalysisService.QueryRuntimeInformation("path: '*$($AFDatabase.name)*' sortBy: 'lastLag' sortOrder: 'Desc'", "id name status");
            #     if ($null -eq $results) {
            #         write-log -v_Message "Pas de resultat sur la requete." -v_ConsoleOutput -v_LogLevel WARN 
            #     }
            #     foreach($result in $results){
            #         $guid = $result[0];
            #         $name = $result[1];
            #         $status = $result[2];
            #         write-log -v_Message "Guid = $guid, Name = $name, Status = $status." -v_ConsoleOutput -v_LogLevel INFO
            #     }
            #     if (($results | ForEach-Object { $_[2] } | Where-Object { $_ -eq "Error" } | Measure-Object).Count -gt 0) {
            #         Write-Log -v_Message "Some analysis listed in the input file are in error, please correct them or remove them from the list." -v_LogLevel ERROR -v_ConsoleOutput
            #         Exit
            #     }
            #     Start-Sleep -Seconds 1
            # }
            # While ($results -and -not ($results | ForEach-Object { $_[2] } | Where-Object { $_ -ne "Running" } | Measure-Object).Count -eq 0)
        }
            
        Write-Log -v_Message "Analysis successfully started." -v_ConsoleOutput -v_LogLevel INFO
        
        # Queue a backfill request to 
        Write-Log -v_Message "Starting Backfill request to the analysis service." -v_ConsoleOutput -v_LogLevel INFO
        $reason = [ref]""
        if ($afAnalysisService.CanQueueCalculation($reason)){
            $QueueCalculationEventID = $afAnalysisService.QueueCalculation($AFAnalysisList, $AFTimeRange, [OSIsoft.AF.Analysis.AFAnalysisService+CalculationMode]::DeleteExistingData)
            Write-Log -v_Message "Calculation started by the analysis service. ID: $QueueCalculationEventID" -v_ConsoleOutput -v_LogLevel INFO
        }
        else {
            Write-Log -v_Message "Calculation cannot be started by the analysis service." -v_ConsoleOutput -v_LogLevel INFO
            throw "`$afAnalysisService.CanQueueCalculation() returned false : $reason"
        }
       
        $BackfillStatus = Wait-EndOfBackfilling -AFAnalysisList $AFAnalysisList -AFTimeRange $AFTimeRange -RecalculationLogFilePath $RecalculationLogFilePath
        if($false -eq $BackfillStatus) { throw "Backfill goes wrong for some analysis."}
    }
    catch {
        Write-Log -v_Message "Failed to backfill analysis: Line $($_.InvocationInfo.ScriptLineNumber) :: $_" -v_LogLevel "ERROR" -v_ConsoleOutput
    }
    finally{
        # Stop the analysis
        Write-Log -v_Message "Stopping all the analysis listed..." -v_ConsoleOutput -v_LogLevel INFO
        [OSIsoft.AF.Analysis.AFAnalysis]::SetStatus($AFAnalysisList, [OSIsoft.AF.Analysis.AFStatus]::Disabled)
        Write-Log -v_Message "Analysis successfully stopped." -v_ConsoleOutput -v_LogLevel INFO
    }
}

function main {
    param(
        [string]$afServerName,
        [string]$afDBName,
        [string]$afSDKPath,
        [string]$InputCsvFilePathAndName,
        [string]$RecalculationLogFilePath,
        [int]$DeltaStartInMinutes = 1440,
        [int]$DeltaEndInMinutes = 0,
        [bool]$AutomaticMode = $true        
    )

    # 00 : PREREQUISITES
    Clear-Host
    Write-Log -v_Message "Script $(Split-Path -Path $MyInvocation.PSCommandPath -Leaf) started" -v_ConsoleOutput -v_LogLevel SUCCESS
    Import-AFSDK -AFSDKPath $afSDKPath
    
    # 01 : CONNECTION TO PI AF AND DATABASE
    $AFDB = Connect-AFServer -afServerName $afServerName -afDBName $afDBName

    # 02 : READ INPUT FILE AND GET THE ANALYSIS
    $AFAnalysisList = Get-AFAnalysisListFromCsvFile -AFDatabase $AFDB -InputCsvFilePathAndName $InputCsvFilePathAndName

    # 03 : CALCULATE THE TIME RANGE OF THE BACKFILL
    $AFTimeRangeToBackfill = Format-AFTimeRange -DeltaStartInMinutes $DeltaStartInMinutes -DeltaEndInMinutes $DeltaEndInMinutes

    # 04 : START THE ANALYSIS, BACKFILL, AND STOP THE ANALYSIS
    Start-AFAnalysisRecalculation -AFDatabase $AFDB -AFAnalysisList $AFAnalysisList -AFTimeRange $AFTimeRangeToBackfill -RecalculationLogFilePath $RecalculationLogFilePath -AutomaticMode $AutomaticMode

    # 05 : DISCONNECT FROM AF SERVER
    Disconnect-AFServer -afServerName $afServerName
    Write-Log -v_Message "Backfilling process successfully completed." -v_ConsoleOutput -v_LogLevel SUCCESS
    if($AutomaticMode -eq $false ) { Read-Host "Press <Enter> to close the program." }
}

# Launch main function
main -afServerName $afServerName -afDBName $afDBName -InputCsvFilePathAndName $InputCsvFilePathAndName -afSDKPath $afSDKPath -DeltaStartInMinutes $DeltaStartInMinutes -DeltaEndInMinutes $DeltaEndInMinutes -RecalculationLogFilePath $RecalculationLogFilePath -AutomaticMode $AutomaticMode