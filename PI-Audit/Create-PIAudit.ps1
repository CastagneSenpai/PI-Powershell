Function Get-PIServerType {
    $ServerTypes = @{}  # Hashtable vide

    # PI Data Archive
    $ServerTypes["PI Data Archive"] = if (Get-Service -ServiceName "piarchss" -ErrorAction SilentlyContinue) { $true } else { $false }

    # PI AF Server
    $ServerTypes["PI Asset Framework"] = if (Get-Service -ServiceName "AFService" -ErrorAction SilentlyContinue) { $true } else { $false }

    # PI Analysis
    $ServerTypes["PI Analysis"] = if (Get-Service -ServiceName "PIAnalysisManager" -ErrorAction SilentlyContinue) { $true } else { $false }

    # Retourner le hashtable contenant les resultats
    return $ServerTypes
}

Function Get-ServerInformation {
    # Recuperer les informations de base du serveur
    $ServeurInformations = @{
        OSVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
        UptimeDays = "{0} days" -f ((Get-Date) - (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime).Days
        CPUTotal = (Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
        MemoryTotalMB = "{0:N2} GB" -f [math]::round((Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    }

    return $ServeurInformations
}

Function Get-PIWindowsServicesInfo {
    $PIWindowsServiceStatus = @{}  # Hashtable vide
    
    # Obtenir les services commençant par "PI"
    Get-Service -DisplayName "PI *" -ErrorAction SilentlyContinue | ForEach-Object {
        # Ajouter chaque service et son statut au hashtable
        $PIWindowsServiceStatus[$_.DisplayName] = ("{0} ({1})" -f $_.Status, $_.StartType)
    }
    
    # Retourner le hashtable
    return $PIWindowsServiceStatus
}

Function Get-PIDATuningParameters{
    param(
        $Con   
    )

    # Creer un hashtable vide
    $tuningParams = @{}

    # Liste des tuning parameters à recuperer (ajouter ceux dont tu as besoin)
    $tuningParamNames = @(
    "Archive_MaxPrivateBytesPercentOfTotal",
    "Archive_MinMemAvail",
    "Archive_LowDiskSpaceMB",
    "Server_AuthenticationPolicy",
    "MaxAuthAttempts",
    "Backup_LowDiskSpaceMB",
    "Replication_ClockDiffLimit",
    "Snapshot_EventQueuePath",
    "Snapshot_AnnotationSizeLimit",
    "Archive_MaxQueryExecutionSec",
    "writeRetry",
    "readRetry",
    "StandAloneMode"
    )

    # Recuperer chaque tuning parameter et l'ajouter dans le hashtable
    foreach ($param in $tuningParamNames) {
        try {
            $result = Get-PITuningParameter -Name $param -Connection $con
            $tuningParams[$param + "_Value"] = if ($result.Value) { $result.Value } else { "N/A" }
            $tuningParams[$param + "_Default"] = if ($result.Default) { $result.Default } else { "N/A" }
        } catch {
            Write-Warning "Erreur lors de la recuperation du paramètre $param"
        }
    }
    return $tuningParams
}  

Function Get-PIDAArchiveDiskInfo{
    param(
        $Con   
    )

    # Hashtable vide
    $driveInfo = @{}

    $ArchivesRootAndFilePrefixe = (Get-PITuningParameter -Name "Archive_AutoArchiveFileRoot" -Connection $con).Value
    $ArchivesDriveLetter = $ArchivesRootAndFilePrefixe.Substring(0,2)
    $drive = Get-PSDrive -Name ((Get-Item $ArchivesDriveLetter).PSDrive.Name)

    # Calculer les espaces disque
    $driveInfo["ArchiveDisk_SpaceUsedGB"] = "{0:N2} GB" -f [math]::Round($drive.Used / 1GB, 2)
    $driveInfo["ArchiveDisk_SpaceFreeGB"] = "{0:N2} GB" -f ([math]::Round($drive.Free / 1GB, 2))
    $driveInfo["ArchiveDisk_PercentUsed"] = "{0:N2} %" -f ([math]::Round(($drive.Used / ($drive.Used + $drive.Free)) * 100, 2))

    Return $driveInfo
}

Function Get-PIDAPointSourcesInfo{
    param(
        $Con   
    )
    
    # Hashtable vide
    $pointSourceInfo = @{}

    $PointSources = Get-PIPointSource -Name "*" -Connection $con
    foreach ($PointSource in $PointSources){
        $pointSourceInfo["$($PointSource.Name)"] = $PointSource.Count
    }
    return $pointSourceInfo 
}

Function Get-PIDADigitalStateInfo{
    param(
        $con
    )
    

    #TODO
    return $null
}

Function Get-PIDAStatistics{
    param(
        $con
    )
    
    $statInfo = @{}
    $stats = Get-PIArchiveStatistics -Connection $con

    ForEach($stat in $stats){
        $statInfo["$($stat.Name)"] = $stat.Value
    }

    return $statInfo
}

Function Get-PIAFInformations{
    param(
        $afServer
    )

    # Initialisation de la variable pour stocker les informations PI AF
    $PIAFInformation = @{}

    # Nombre de bases de donnees AF sur le serveur
    $afDatabases = $afServer.Databases
    $PIAFInformation["AFDatabasesCount"] = $afDatabases.Count

    # Parcourir chaque base de donnees pour recuperer les informations detaillees
    foreach ($afDatabase in $afDatabases) {
        # Nom de la base de donnees
        $dbName = $afDatabase.Name

        # Nombre total d'elements
        $PIAFInformation[$dbName + "_NombreTotalElements"] = ($afDatabase.Elements | ForEach-Object { $_; if ($_.Elements.Count -gt 0) { $_.Elements | ForEach-Object { $_; if ($_.Elements.Count -gt 0) { $_.Elements } } } }).Count

        # Nombre total de templates
        $PIAFInformation[$dbName + "_NombreTotalTemplates"] = $afDatabase.ElementTemplates.Count

        # Ratio d'elements bases sur des templates
        $templatedElementsCount = ($afDatabase.Elements | ForEach-Object { $_.Elements.Count; $_.Elements | ForEach-Object { $null -ne $_.Template} }).Count

        if ($afDatabase.Elements.Count -gt 0) {
            $PIAFInformation[$dbName + "_RatioElementsTempletises"] = ([math]::Round(($templatedElementsCount / $PIAFInformation[$dbName + "_NombreTotalElements"]) * 100, 2)).ToString() + "%"
        } else {
            $PIAFInformation[$dbName + "_RatioElementsTempletises"] = 0
        }

        # Nombre total de tables
        $PIAFInformation[$dbName + "_NombreTotalTables"] = $afDatabase.Tables.Count

        # Nombre total de template d'analyses
        $PIAFInformation[$dbName + "_NombreTotalTemplatesAnalyses"] = $afDatabase.AnalysisTemplates.Count
    }
    $afServer.Disconnect()
    return $PIAFInformation
}

Function Show-Report {
    param(
        $AuditReport
    )
    Write-Output "--- AUDIT SERVER $env:COMPUTERNAME --- "

    # Parcourir chaque section dans le rapport
    foreach ($SectionName in $AuditReport.PSObject.Properties.Name) {
        # Afficher le titre de la section
        Write-Output "Information de $SectionName :"
        
        # Recuperer les elements de la section et appliquer l'indentation
        if($AuditReport.$SectionName){
            $AuditReport.$SectionName.GetEnumerator() | Sort-Object Key |  ForEach-Object {
                Write-Output ("  - {0} : {1}" -f $_.Key, $_.Value)
            }
        }
        Write-Output "" # Ligne vide pour separer les sections
    }
}

Function Write-ReportToFile{
    param(
        [Parameter(Mandatory=$true)]
        $AuditReport,
        
        [Parameter(Mandatory=$false)]
        [string]$OutputReportFile = (join-path $PSScriptRoot ("Audit-$env:COMPUTERNAME-" + (Get-Date -format "yyyyMMddThhmm") + ".log")),

        [Parameter(Mandatory=$false)]
        [string]$OutputCSVFile = (join-path $PSScriptRoot ("Audit-$env:COMPUTERNAME-" + (Get-Date -format "yyyyMMddThhmm") + ".csv"))
    )

    # ecrire l'en-tête d'audit dans le fichier Report
    Add-Content -Path $OutputReportFile -Value "--- AUDIT SERVER $env:COMPUTERNAME ---" 

    # Creer une liste pour stocker les lignes CSV
    $csvLines = @()
    $csvLines += "Parametre;Valeur"

    # Parcourir chaque section dans le rapport
    foreach ($SectionName in $AuditReport.PSObject.Properties.Name) {
        # ecrire le titre de la section dans le fichier
        Add-Content -Path $OutputReportFile -Value "Information de $SectionName :"
        
        # Recuperer les elements de la section et appliquer l'indentation
        if ($AuditReport.$SectionName) {
            $AuditReport.$SectionName.GetEnumerator() | Sort-Object Key | ForEach-Object {
                Add-Content -Path $OutputReportFile -Value ("  - {0} : {1}" -f $_.Key, $_.Value) # MAJ Rapport
                $csvLines += "{0};{1}" -f $_.Key, $_.Value # MAJ CSV
            }
        }
        # Ligne vide pour separer les sections
        Add-Content -Path $OutputReportFile -Value ""
    }
    $csvLines | Out-File -FilePath $OutputCSVFile -Encoding UTF8
}

Function Main {
    Clear-Host

    $ServerType = Get-PIServerType
    $ServerInformations = Get-ServerInformation
    $PIWindowsServiceStatus = Get-PIWindowsServicesInfo

    if($ServerType['PI Data Archive'])  { 
        # Connexion au PI Data Archive
        import-module (Join-Path $PSScriptRoot '..\lib\connection.psm1')
        $con = Connect-PIDataArchive -PIDataArchiveMachineName "$env:COMPUTERNAME"
        
        $PIDATuningParameters = Get-PIDATuningParameters -con $con
        $PIDAArchivesDiskInfo = Get-PIDAArchiveDiskInfo -con $con
        $PIDAPointSourcesInfo = Get-PIDAPointSourcesInfo -Con $con
        $PIDADigitalStates = Get-PIDADigitalStateInfo -con $con
        $PIDAStatisticsInfo = Get-PIDAStatistics -con $con
    }

    if($ServerType['PI Asset Framework'])   {
        # Creation de l'objet PI Systems et connexion au serveur
        $afSystems = New-Object OSIsoft.AF.PISystems
        $afServer = $afSystems[$env:COMPUTERNAME]
        $afServer.Connect()

        $PIAFInformations = Get-PIAFInformations -afServer $afServer
    }

    # Fusionner les resultats en un seul objet
    $AuditReport = [PSCustomObject]@{
        Server_Informations = $ServerInformations
        Windows_Service_Status = $PIWindowsServiceStatus
        PI_DA_Statistics = $PIDAStatisticsInfo
        PI_DA_Tuning_Parameters = $PIDATuningParameters
        PI_DA_ArchivesDisk = $PIDAArchivesDiskInfo
        PI_DA_Point_Sources = $PIDAPointSourcesInfo
        PI_DA_Digital_States = $PIDADigitalStates
        PI_AF_Informations = $PIAFInformations
    }

    # Gestion du rapport - console + fichier
    Show-Report -AuditReport $AuditReport
    Write-ReportToFile -AuditReport $AuditReport
}

# APPEL FONCTION MAIN
Main