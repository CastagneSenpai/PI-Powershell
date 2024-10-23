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
        [string]$afServerName = "ACEW1DSTEKPIS01",
        [string]$afDBName = "Test_Prd_Posting",
        [string]$afSDKPath = "D:\OSISOFT\PIPC_x86\AF\PublicAssemblies\4.0\OSIsoft.AFSDK.dll",
        [string]$InputCsvFilePathAndName = (Join-path $PSScriptRoot "input.csv"),
        [int]$DeltaStartInMinutes = 1440,
        [int]$DeltaEndInMinutes = 30,
        [bool]$AutomaticMode = $false
)

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

# Fonction d'ajout de la librairie AFSDK si non ajoutÃ©e prÃ©cÃ©demment 
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

# Fonction de connection Ã  PI AF, retourne l'objet de la connection de type [OSIsoft.AF.AFDatabase]
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

# Fonction qui retourne une liste d'objet AFAnalysis Ã  partir d'un fichier CSV
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

# Fonction de construction de l'objet AFTimeRange Ã  partir d'un offset et d'une durÃ©e de backfill souhaitÃ©e
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

function Start-VisualSleep {
    param(
        [Int] $Seconds = 30,
        [String] $Activity = "Attente en cours"
    )

    # Créer une nouvelle fenêtre
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Activity
    $form.Size = New-Object System.Drawing.Size(300, 100)
    
    # Créer une barre de progression
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 20)
    $progressBar.Size = New-Object System.Drawing.Size(250, 30)
    $progressBar.Style = 'Continuous'
    
    # Ajouter la barre de progression à la fenêtre
    $form.Controls.Add($progressBar)
    
    # Afficher la fenêtre
    $form.Show()

    # Définir l'intervalle en millisecondes pour une mise à jour plus fluide
    $interval = 100  # En millisecondes (0.1 seconde)
    $totalIterations = ($Seconds * 1000) / $interval  # Nombre total d'itérations

    # Démarrer la mise à jour de la barre de progression
    for ($i = 1; $i -le $totalIterations; $i++) {
        $progressBar.Value = ($i / $totalIterations) * 100  # Mettre à jour la valeur de la barre de progression
        Start-Sleep -Milliseconds $interval  # Pause entre chaque mise à jour
    }

    # Fermer la fenêtre après l'attente
    $form.Close()
}

# Fonction de lancement d'un backfill de plusieurs analyses sur une AFTimeRange
function Start-AFAnalysisRecalculation{
    param(
        $AFDatabase,
        [System.Collections.Generic.List[OSIsoft.AF.Analysis.AFAnalysis]] $AFAnalysisList,
        [OSIsoft.AF.Time.AFTimeRange] $AFTimeRange,
        $AutomaticMode
    )

    $afAnalysisService = $AFDatabase.PISystem.AnalysisService
    try {  
        if($null -eq $AFAnalysisList){
            throw "There is no valid AFAnalysis in the input CSV file."
        }
        # Start AF Analysis
        Write-Log -v_Message "Starting all the analysis listed..." -v_ConsoleOutput -v_LogLevel INFO
        [OSIsoft.AF.Analysis.AFAnalysis]::SetStatus($AFAnalysisList, [OSIsoft.AF.Analysis.AFStatus]::Enabled)

        # TODO : Evaluer le status des analyses avant de continuer
        if($AutomaticMode -eq $true) { Start-VisualSleep -Seconds 12 -Activity "Analysis starting..." }
        if($AutomaticMode -eq $false ) { Read-Host "Pause - Validate Analysis started." }
        # METHODE 1 : GetStatus -- KO : retourne vide
            # $statuses = [OSIsoft.AF.Analysis.AFAnalysis]::GetStatus($AFAnalysisList)
            # Write-Log -v_Message "Status: $statuses" -v_ConsoleOutput
        #METHODE 2 : QueryRuntimeInformation -- KO : Status passe de Stopped Ã  Running sans passer par Starting.
        #    Do{
        #        $results = $afAnalysisService.QueryRuntimeInformation("path: '*$($AFDatabase.name)*' sortBy: 'lastLag' sortOrder: 'Desc'", "id name status");
        #        if ($null -eq $results) {
        #            write-log -v_Message "Pas de resultat sur la requete." -v_ConsoleOutput -v_LogLevel WARN 
        #        }
        #        foreach($result in $results){
        #            $guid = $result[0];
        #            $name = $result[1];
        #            $status = $result[2];
        #            write-log -v_Message "Guid = $guid, Name = $name, Status = $status." -v_ConsoleOutput -v_LogLevel INFO
        #        }
        #    }
        #    While ($results[0].status -ne "Running")        
            
        Write-Log -v_Message "Analysis successfully started." -v_ConsoleOutput -v_LogLevel INFO
        
        # Queue a backfill request to 
        Write-Log -v_Message "Starting Backfill request to the analysis service." -v_ConsoleOutput -v_LogLevel INFO
        $QueueCalculationEventID = $afAnalysisService.QueueCalculation($AFAnalysisList, $AFTimeRange, [OSIsoft.AF.Analysis.AFAnalysisService+CalculationMode]::DeleteExistingData)
        Write-Log -v_Message "Calculation started by the analysis service. ID: $QueueCalculationEventID" -v_ConsoleOutput -v_LogLevel INFO

        # TODO : Evaluer le temps de la CalculationQueue
            # MÃ©thode 1 : QueryRuntimeInformation -- KO car Guid retournÃ© concerne les analyses et non le backfilling lancÃ©.
        if($AutomaticMode -eq $true) { Start-VisualSleep -Seconds 15 -Activity "Calculation Queue in progress ..." }
        if($AutomaticMode -eq $false ) { Read-Host "Pause - Validate backfill OK" }
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
        [string]$afServerName,
        [string]$afDBName,
        [string]$afSDKPath,
        [string]$InputCsvFilePathAndName,
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
    Start-AFAnalysisRecalculation -AFDatabase $AFDB -AFAnalysisList $AFAnalysisList -AFTimeRange $AFTimeRangeToBackfill -AutomaticMode $AutomaticMode

    # 05 : DISCONNECT FROM AF SERVER
    Disconnect-AFServer -afServerName $afServerName
    Write-Log -v_Message "Backfilling process successfully completed." -v_ConsoleOutput -v_LogLevel SUCCESS
    if($AutomaticMode -eq $false ) { Read-Host "Press <Enter> to close the program." }
}

# Lancement fonction principale
main -afServerName $afServerName -afDBName $afDBName -InputCsvFilePathAndName $InputCsvFilePathAndName -afSDKPath $afSDKPath -DeltaStartInMinutes $DeltaStartInMinutes -DeltaEndInMinutes $DeltaEndInMinutes -AutomaticMode $AutomaticMode