<#
.SYNOPSIS
    Script to manage AFAnalysis backfill processes and PI AF analysis operations.

.DESCRIPTION
    This script automates the backfilling of AFAnalysis in PI AF using OSIsoft AF SDK. It includes connection management to AF servers, CSV-based input for analyses, log handling, and configurable time ranges for the backfill. 
    It also handles the start and stop of the specified analyses.

.NOTES
    Developer  : Romain Castagne
    Email      : romain.castagne-ext@syensqo.com
    Company    : CGI - SYENSQO
    Date       : 23/10/2024

    Version 1.0 : Script start the analysis, user manually validate that the analysis are running ; Same for backfilling process.
    Version 1.1 : Script use categories to filter analysis to backfill, it can process multiple cycle (OEE before EF ...)
    Version 2.0 : Script detect automatically when the analysis are started, and when the backfill is finished using the analysis log file.
    Version 2.1 : User can choose to start/stop the analysis, or let them in start mode.
    Version 2.2 : Task sent with .bat that specify all the parameters, default values not available anymore to keep one script.
    Version 2.3 : Use StartTime instead of EndTime to match the analysis log (for some EF, endTime <> logLine associated).

.PARAMETER afServerName
    The name of the PI AF server to connect to.

.PARAMETER afDBName
    The name of the PI AF database to use.

.PARAMETER afSDKPath
    Path to the OSIsoft AF SDK DLL.

.PARAMETER DeltaStartInMinutes
    Time range start offset in minutes from the current time (for backfill).

.PARAMETER DeltaEndInMinutes
    Time range end offset in minutes from the current time (for backfill).

.PARAMETER CategoriesName
    List of analysis categories name that will be included for the backfilling. Order of categories on the list matters for analysis dependencies. 

.PARAMETER AutomaticMode
    Set to true for automatic pauses, false for manual validation.

#>


param(
        [string]$afServerName,
        [string]$afDBName,
        [string]$afSDKPath,
        [string]$RecalculationLogFilePath,
        [int]$DeltaStartInMinutes,
        [int]$DeltaEndInMinutes,
        [string[]]$CategoriesName,
        [bool]$AutomaticMode,
        [bool]$StartAndStopAnalysis
)

# Log management function
function Write-Log {
    [CmdletBinding()]
    Param(
        [string]$v_Message,
        [string]$v_LogFile = (Join-Path -Path $PSScriptRoot -ChildPath ("Logs\" + (Get-Date -Format yyyy-MM-dd) + "_Logs.txt")),
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
        [string]$afServerName,
        [string]$afDBName
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

# Function that returns a list of AFAnalysis based on an analysis category
function Get-AFAnalysisListByCategory {
    param(
        $AFDatabase,
        [string] $CategoryName = "Auto-Backfill"
    )

    $AnalysisList = [System.Collections.Generic.List[OSIsoft.AF.Analysis.AFAnalysis]]::new()

    foreach($CurrentAnalysis in $AFDatabase.analyses){
        foreach($AnalysisCategory in $CurrentAnalysis.Categories){
            if($AnalysisCategory.name -eq $CategoryName){
                $AnalysisList.Add($CurrentAnalysis)
                break
            }
        }
    }

    return $AnalysisList
}

# Function for constructing the AFTimeRange object from an offset and a desired backfill duration
function Format-AFTimeRange{
    param(
        [Int] $DeltaStartInMinutes,
        [Int] $DeltaEndInMinutes
    )

    $endTime = New-Object OSIsoft.AF.Time.AFTime((Get-Date).AddMinutes((-$DeltaEndInMinutes)))
    $startTime = New-Object OSIsoft.AF.Time.AFTime((Get-Date).AddMinutes(-$DeltaStartInMinutes))
    $afTimeRange = New-Object OSIsoft.AF.Time.AFTimeRange($startTime, $endTime)
    
    Write-Log -v_Message "The backfill time range is set to start at $($startTime) and end at $($endTime)." -v_ConsoleOutput -v_LogLevel INFO
    return $afTimeRange
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
 #   $AFTimeRangeConvertedToMatchLogFile = $AFTimeRange.StartTime.ToString("yyyy-MM-ddThh:00:00") + "--" + $AFTimeRange.EndTime.ToString("yyyy-MM-ddThh:00:00")
    $AFTimeRangeConvertedToMatchLogFile = $AFTimeRange.StartTime.ToString("yyyy-MM-ddThh:00:00")
 
    # Loop to wait for backfilling to finish
    $nbTry=0
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
                # $LogDateFin = (Get-Date $LogDateArray[1]).ToString("yyyy-MM-ddThh:00:00")
                # $LogLineTimeRangeFormatted = $LogDateDebut + "--" + $LogDateFin
    
                # Check that the log corresponds to the analysis and the specified period
                # if ($logLine.Id -eq $AnalysisWithStatus.Analysis.UniqueID -and $LogLineTimeRangeFormatted -eq $AFTimeRangeConvertedToMatchLogFile -and $logLine.Type -eq "Manual") {
                if ($logLine.Id -eq $AnalysisWithStatus.Analysis.UniqueID -and $LogDateDebut -eq $AFTimeRangeConvertedToMatchLogFile -and $logLine.Type -eq "Manual") {
    
                    # Case when status is Completed
                    if ($logLine.Status -in ("Completed", "PendingCompletion")) {
                        Write-Log -v_Message "--- Backfilling of analysis $($AnalysisWithStatus.analysis.target)[$($AnalysisWithStatus.analysis.name)] completed." -v_LogLevel DEBUG -v_ConsoleOutput
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
        }
    
        # Pause before next check
        Start-Sleep -Seconds 10
        $nbTry++
    
    } while ($AnalysisWithStatusList.Status -contains "InProgress" -and -not ($AnalysisWithStatusList.Status -contains "Error") -and $nbTry -ile 15)
    if ($AnalysisWithStatusList.Status -contains "Error") {
        return $false
    }
    else{
        return $true
    }
}

# Function that start analysis, wait for Running status, launch a backfill based on the AFTimeRange set, wait the end of backfilling, and stop the analysis
function Start-AFAnalysisRecalculation{
    param(
        $AFDatabase,
        [System.Collections.Generic.List[OSIsoft.AF.Analysis.AFAnalysis]] $AFAnalysisList,
        [OSIsoft.AF.Time.AFTimeRange] $AFTimeRange,
        [bool]$AutomaticMode,
        [string] $RecalculationLogFilePath,
        [string]$CategoryName,
        [bool]$StartAndStopAnalysis
    )

    $afAnalysisService = $AFDatabase.PISystem.AnalysisService
    try {  
        if($null -eq $AFAnalysisList){
            throw "There is no AFAnalysis with the category Auto-Backfill to process."
        }

        # Start AF Analysis
        if($StartAndStopAnalysis){
            Write-Log -v_Message "Starting all the analysis listed..." -v_ConsoleOutput -v_LogLevel INFO
            [OSIsoft.AF.Analysis.AFAnalysis]::SetStatus($AFAnalysisList, [OSIsoft.AF.Analysis.AFStatus]::Enabled)
    
            # Wait that analysis are well started
            if($AutomaticMode -eq $false ) { Read-Host "Pause - Validate Analysis started." }
            if($AutomaticMode -eq $true) { 
                
                # QueryRuntimeInformation -- Apply a filter by Category name = Auto-Backfill
                Do{
                    $results = $afAnalysisService.QueryRuntimeInformation("path: '*$($AFDatabase.name)*' Category: '$CategoryName' sortBy: 'lastLag' sortOrder: 'Desc'", "id name status");
                    if ($null -eq $results) {
                        write-log -v_Message "Pas de resultat sur la requete." -v_ConsoleOutput -v_LogLevel WARN 
                    }
                    foreach($result in $results){
                        $name = $result[1];
                        $status = $result[2];
                        Write-log -v_Message "Name = $name, Status = $status." -v_ConsoleOutput -v_LogLevel DEBUG
                    }
                    if (($results | ForEach-Object { $_[2] } | Where-Object { $_ -eq "Error" } | Measure-Object).Count -gt 0) {
                        Write-Log -v_Message "Some analysis listed in the input file are in error, please correct them or remove them from the list." -v_LogLevel ERROR -v_ConsoleOutput
                        Exit
                    }
                    Start-Sleep -Seconds 1
                }
                While ($results -and -not ($results | ForEach-Object { $_[2] } | Where-Object { $_ -ne "Running" } | Measure-Object).Count -eq 0)
            }
                
            Write-Log -v_Message "Analysis successfully started." -v_ConsoleOutput -v_LogLevel INFO
        }
        else {
            Write-Log -v_Message "The parameter 'StartAndStopAnalysis' is defined as 'false', analysis must be running." -v_ConsoleOutput -v_LogLevel INFO
        }
        
        
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
       
        if($AutomaticMode -eq $true) { $BackfillStatus = Wait-EndOfBackfilling -AFAnalysisList $AFAnalysisList -AFTimeRange $AFTimeRange -RecalculationLogFilePath $RecalculationLogFilePath }
        if($AutomaticMode -eq $false ) { Read-Host "Pause - Validate backfilling is over." }
        if($false -eq $BackfillStatus) { throw "Backfill goes wrong for some analysis."}
    }
    catch {
        Write-Log -v_Message "Failed to backfill analysis: Line $($_.InvocationInfo.ScriptLineNumber) :: $_" -v_LogLevel "ERROR" -v_ConsoleOutput
    }
    finally{
        if ($StartAndStopAnalysis){
            # Stop the analysis
            Write-Log -v_Message "Stopping all the analysis listed..." -v_ConsoleOutput -v_LogLevel INFO
            [OSIsoft.AF.Analysis.AFAnalysis]::SetStatus($AFAnalysisList, [OSIsoft.AF.Analysis.AFStatus]::Disabled)
            Write-Log -v_Message "Analysis successfully stopped." -v_ConsoleOutput -v_LogLevel INFO
        }
        else{
            Write-Log -v_Message "The parameter 'StartAndStopAnalysis' is defined as 'false', the script let the analysis running." -v_ConsoleOutput -v_LogLevel INFO
        }
        
    }
}

function main {
    param(
        [string]$afServerName,
        [string]$afDBName,
        [string]$afSDKPath,
        [string]$InputCsvFilePathAndName,
        [string]$RecalculationLogFilePath,
        [int]$DeltaStartInMinutes,
        [int]$DeltaEndInMinutes,
        [string[]]$CategoriesName,
        [bool]$AutomaticMode,
        [bool]$StartAndStopAnalysis
    )

    # 00 : PREREQUISITES
    Clear-Host
    Write-Log -v_Message "Script $(Split-Path -Path $MyInvocation.PSCommandPath -Leaf) started" -v_ConsoleOutput -v_LogLevel INFO
    Write-Log -v_Message "--- Parameter afServerName = $afServerName" -v_LogLevel DEBUG -v_ConsoleOutput
    Write-Log -v_Message "--- Parameter afDBName = $afDBName" -v_LogLevel DEBUG -v_ConsoleOutput
    Write-Log -v_Message "--- Parameter afSDKPath = $afSDKPath" -v_LogLevel DEBUG -v_ConsoleOutput
    Write-Log -v_Message "--- Parameter RecalculationLogFilePath = $RecalculationLogFilePath" -v_LogLevel DEBUG -v_ConsoleOutput
    Write-Log -v_Message "--- Parameter DeltaStartInMinutes = $DeltaStartInMinutes" -v_LogLevel DEBUG -v_ConsoleOutput
    Write-Log -v_Message "--- Parameter DeltaEndInMinutes = $DeltaEndInMinutes" -v_LogLevel DEBUG -v_ConsoleOutput
    Write-Log -v_Message "--- Parameter CategoriesName = $CategoriesName" -v_LogLevel DEBUG -v_ConsoleOutput
    Write-Log -v_Message "--- Parameter AutomaticMode = $AutomaticMode" -v_LogLevel DEBUG -v_ConsoleOutput
    Write-Log -v_Message "--- Parameter StartAndStopAnalysis = $StartAndStopAnalysis" -v_LogLevel DEBUG -v_ConsoleOutput
    
    # 01 : CONNECTION TO PI AF AND DATABASE
    Import-AFSDK -AFSDKPath $afSDKPath
    $AFDB = Connect-AFServer -afServerName $afServerName -afDBName $afDBName

    # 02 : CALCULATE THE TIME RANGE OF THE BACKFILL
    $AFTimeRangeToBackfill = Format-AFTimeRange -DeltaStartInMinutes $DeltaStartInMinutes -DeltaEndInMinutes $DeltaEndInMinutes

    foreach($CategoryName in $CategoriesName)
    {
        Write-Log -v_Message "Processing category <$CategoryName>. " -v_ConsoleOutput -v_LogLevel INFO
        # 03 : GET THE ANALYSIS BASED ON CATEGORY NAME
        $AFAnalysisList = Get-AFAnalysisListByCategory -AFDatabase $AFDB -CategoryName $CategoryName

        # 04 : START THE ANALYSIS, BACKFILL, AND STOP THE ANALYSIS
        Start-AFAnalysisRecalculation -AFDatabase $AFDB -AFAnalysisList $AFAnalysisList -AFTimeRange $AFTimeRangeToBackfill -RecalculationLogFilePath $RecalculationLogFilePath -CategoryName $CategoryName -AutomaticMode $AutomaticMode -StartAndStopAnalysis $StartAndStopAnalysis
        Write-Log -v_Message "Category <$CategoryName> processed. " -v_ConsoleOutput -v_LogLevel INFO
    }
    
    # 05 : DISCONNECT FROM AF SERVER
    Disconnect-AFServer -afServerName $afServerName
    Write-Log -v_Message "Backfilling process successfully completed." -v_ConsoleOutput -v_LogLevel SUCCESS
    if($AutomaticMode -eq $false ) { Read-Host "Press <Enter> to close the program." }
}

# Launch main function
main -afServerName $afServerName -afDBName $afDBName -afSDKPath $afSDKPath -DeltaStartInMinutes $DeltaStartInMinutes -DeltaEndInMinutes $DeltaEndInMinutes -RecalculationLogFilePath $RecalculationLogFilePath -CategoriesName $CategoriesName -AutomaticMode $AutomaticMode -StartAndStopAnalysis $StartAndStopAnalysis