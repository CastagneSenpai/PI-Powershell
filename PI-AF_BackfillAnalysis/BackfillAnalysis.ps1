# Fonction de gestion des logs 
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
                    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                    Write-Host $logEntry -ForegroundColor $color
                }
            }
            catch {
                Write-Error "Failed to write log: $_"
            }
        }
    }
}

# Fonction d'ajout de la librairie AFSDK si non ajoutée précédemment 
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

# Fonction de connection à PI AF, retourne l'objet de la connection de type [OSIsoft.AF.AFDatabase]
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

# Fonction qui retourne une liste d'objet AFAnalysis à partir d'un fichier CSV
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
            Write-Log -v_ConsoleOutput -v_Message "Analysis `'$($row.ElementPath)\$AFAnalysisToAdd`' added for backfilling process."
        }
    }
    return $AnalysisList
}

# Fonction de construction de l'objet AFTimeRange à partir d'un offset et d'une durée de backfill souhaitée
function Format-AFTimeRange{
    param(
        [Int] $BackfillDurationInMinutes = 1440, # A day per default
        [Int] $OffsetInMinutes = 30             # 30 minutes offset
    )

    $endTime = New-Object OSIsoft.AF.Time.AFTime((Get-Date).AddMinutes((-$OffsetInMinutes)))
    $startTime = New-Object OSIsoft.AF.Time.AFTime((Get-Date).AddMinutes(-$BackfillDurationInMinutes))
    $afTimeRange = New-Object OSIsoft.AF.Time.AFTimeRange($startTime, $endTime)
    
    Write-Log -v_Message "The backfill time range is set to start at $($startTime) and end at $($endTime)." -v_ConsoleOutput -v_LogLevel INFO
    return $afTimeRange
}

# Fonction de lancement d'un backfill de plusieurs analyses sur une AFTimeRange
function Start-AFAnalysisRecalculation{
    param(
        $AFDatabase,
        [System.Collections.Generic.List[OSIsoft.AF.Analysis.AFAnalysis]] $AFAnalysisList,
        [OSIsoft.AF.Time.AFTimeRange] $AFTimeRange
    )

    try {  
        if($null -eq $AFAnalysisList){
            throw "There is no valid AFAnalysis in the input CSV file."
        }
        # Start AF Analysis
        Write-Log -v_Message "Starting all the analysis listed..." -v_ConsoleOutput -v_LogLevel INFO
        [OSIsoft.AF.Analysis.AFAnalysis]::SetStatus($AFAnalysisList, [OSIsoft.AF.Analysis.AFStatus]::Enabled)
    
        # TODO : Evaluer le status des analyses avant de continuer
        # $statuses = [OSIsoft.AF.Analysis.AFAnalysis]::GetStatus($AFAnalysisList)
        # Write-Log -v_Message "Type of returned data: $($statuses.GetType().FullName)" -v_ConsoleOutput
        # Write-Log -v_Message "Status: $statuses" -v_ConsoleOutput
        Start-Sleep -Seconds 10
        Write-Log -v_Message "Analysis successfully started." -v_ConsoleOutput -v_LogLevel INFO

        # Queue a backfill request to 
        Write-Log -v_Message "Starting Backfill request to the analysis service." -v_ConsoleOutput -v_LogLevel INFO
        $afAnalysisService = $AFDatabase.PISystem.AnalysisService
        $QueueCalculationEventID = $afAnalysisService.QueueCalculation($AFAnalysisList, $AFTimeRange, [OSIsoft.AF.Analysis.AFAnalysisService+CalculationMode]::DeleteExistingData)
        Write-Log -v_Message "Calculation started by the analysis service. ID: $QueueCalculationEventID" -v_ConsoleOutput -v_LogLevel INFO
        Start-Sleep -Seconds 10
        Write-Log -v_Message "Calculation successfully finished." -v_ConsoleOutput -v_LogLevel INFO

        Write-Log -v_Message "Stopping all the analysis listed..." -v_ConsoleOutput -v_LogLevel INFO
        [OSIsoft.AF.Analysis.AFAnalysis]::SetStatus($AFAnalysisList, [OSIsoft.AF.Analysis.AFStatus]::Disabled)
        Write-Log -v_Message "Analysis successfully stopped." -v_ConsoleOutput -v_LogLevel INFO
    }
    catch {
        Write-Log -v_Message "Failed to backfill analysis: Line $($_.InvocationInfo.ScriptLineNumber) :: $_" -v_LogLevel "ERROR" -v_ConsoleOutput
    }
}

function main {
    param(
        [string]$afServerName = "vmcegdidev001",
        [string]$afDBName = "Romain_Dev",
        [string]$afSDKPath = "E:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0\OSIsoft.AFSDK.dll",
        [string]$InputCsvFilePathAndName = (Join-path $PSScriptRoot "Input.csv")
    )

    # 00 : PREREQUISITES
    Clear-Host
    Write-Log -v_Message "Script $(Split-Path -Path $MyInvocation.PSCommandPath -Leaf) started" -v_ConsoleOutput -v_LogLevel SUCCESS
    Import-AFSDK -AFSDKPath $afSDKPath    
    
    # 01 : CONNECTION 
    $AFDB = Connect-AFServer -afServerName $afServerName -afDBName $afDBName

    # 02 : READ INPUT FILE AND GET THE ANALYSIS
    $AFAnalysisList = Get-AFAnalysisListFromCsvFile -AFDatabase $AFDB -InputCsvFilePathAndName $InputCsvFilePathAndName

    # 03 : CALCULATE THE TIME RANGE OF THE BACKFILL
    $AFTimeRangeToBackfill = Format-AFTimeRange -BackfillDurationInMinutes 1440 -OffsetInMinutes 30

    # 04 : START THE ANALYSIS, BACKFILL, AND STOP THE ANALYSIS
    Start-AFAnalysisRecalculation -AFDatabase $AFDB -AFAnalysisList $AFAnalysisList -AFTimeRange $AFTimeRangeToBackfill

    # 05 : DISCONNECT FROM AF SERVER
    Disconnect-AFServer -afServerName $afServerName
    Write-Log -v_Message "Backfilling process successfully completed." -v_ConsoleOutput -v_LogLevel SUCCESS
}

# Lancement fonction principale
main -afServerName "vmcegdidev001" -afDBName "Romain_Dev" -InputCsvFilePathAndName (Join-path $PSScriptRoot "input.csv") -afSDKPath "E:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0\OSIsoft.AFSDK.dll"