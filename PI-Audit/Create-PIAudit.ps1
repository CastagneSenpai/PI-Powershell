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

Function Show-FormattedReportGUI {
    param(
        [string]$csvFilePath = (join-path $PSScriptRoot ("Audit-$env:COMPUTERNAME-" + (Get-Date -format "yyyyMMddThhmm") + ".csv"))
    )

    # Charger les assemblys nécessaires pour les graphiques
    Add-Type -AssemblyName "System.Windows.Forms"
    Add-Type -AssemblyName "System.Windows.Forms.DataVisualization"

    # Lire le fichier CSV
    $data = Import-Csv -Path $csvFilePath -Delimiter ';'

    # Créer la fenêtre principale
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Rapport de l'état du serveur"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = 'CenterScreen'

    # Ajouter une barre de titre (Label)
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Rapport du serveur - " + $env:COMPUTERNAME
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.AutoSize = $true
    $titleLabel.Location = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($titleLabel)

    # Ajouter un graphique circulaire pour les ratios d'éléments templatés
    $chart = New-Object Windows.Forms.DataVisualization.Charting.Chart
    $chart.Width = 300
    $chart.Height = 300
    $chartArea = New-Object Windows.Forms.DataVisualization.Charting.ChartArea
    $chart.ChartAreas.Add($chartArea)
    
    $series = New-Object Windows.Forms.DataVisualization.Charting.Series
    $series.Name = "Templates"
    $series.Points.AddXY("Templated", [double]($data | Where-Object { $_.Parametre -like '*RatioElementsTempletises*' }).Valeur.Split('%')[0])
    $series.Points.AddXY("Non-Templated", 100 - [double]($data | Where-Object { $_.Parametre -like '*RatioElementsTempletises*' }).Valeur.Split('%')[0])
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
    $chart.Series.Add($series)

    $chart.Location = New-Object System.Drawing.Point(10, 50)
    $form.Controls.Add($chart)

    # Ajouter un Label pour l'état des services PI
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Services PI"
    $statusLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $statusLabel.AutoSize = $true
    $statusLabel.Location = New-Object System.Drawing.Point(350, 10)
    $form.Controls.Add($statusLabel)

    # Affichage des services PI avec couleur d'état
    $yPos = 50
    $data | Where-Object { $_.Parametre -like 'PI*Service' } | ForEach-Object {
        $serviceStatusLabel = New-Object System.Windows.Forms.Label
        $serviceStatusLabel.Text = "$($_.Parametre): $($_.Valeur)"
        $serviceStatusLabel.AutoSize = $true
        $serviceStatusLabel.Location = New-Object System.Drawing.Point(350, $yPos)

        # Appliquer des couleurs en fonction de l'état du service
        if ($_.Valeur -like "Running*") {
            $serviceStatusLabel.ForeColor = [System.Drawing.Color]::Green
        } elseif ($_.Valeur -like "Stopped*") {
            $serviceStatusLabel.ForeColor = [System.Drawing.Color]::Red
        } else {
            $serviceStatusLabel.ForeColor = [System.Drawing.Color]::Orange
        }

        $form.Controls.Add($serviceStatusLabel)
        $yPos += 30
    }

    # Ajouter des barres de progression pour l'usage de la mémoire
    $memoryLabel = New-Object System.Windows.Forms.Label
    $memoryLabel.Text = "Mémoire Totale (MB)"
    $memoryLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $memoryLabel.AutoSize = $true
    $memoryLabel.Location = New-Object System.Drawing.Point(10, 380)
    $form.Controls.Add($memoryLabel)

    $memoryBar = New-Object System.Windows.Forms.ProgressBar
    $memoryBar.Location = New-Object System.Drawing.Point(10, 410)
    $memoryBar.Width = 300
    $memoryBar.Height = 30
    $memoryBar.Maximum = [int]($data | Where-Object { $_.Parametre -eq "MemoryTotalMB" }).Valeur
    $memoryBar.Value = [int]($data | Where-Object { $_.Parametre -eq "MemoryTotalMB" }).Valeur
    $form.Controls.Add($memoryBar)

    # Affichage des valeurs de configuration
    $configLabel = New-Object System.Windows.Forms.Label
    $configLabel.Text = "Configuration des Bases PI AF"
    $configLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $configLabel.AutoSize = $true
    $configLabel.Location = New-Object System.Drawing.Point(350, $yPos + 20)
    $form.Controls.Add($configLabel)

    $yPos += 50
    $data | Where-Object { $_.Parametre -like 'AFDatabasesCount' -or $_.Parametre -like '*NombreTotal*' } | ForEach-Object {
        $configStatusLabel = New-Object System.Windows.Forms.Label
        $configStatusLabel.Text = "$($_.Parametre): $($_.Valeur)"
        $configStatusLabel.AutoSize = $true
        $configStatusLabel.Location = New-Object System.Drawing.Point(350, $yPos)

        $form.Controls.Add($configStatusLabel)
        $yPos += 30
    }

    # Afficher la fenêtre
    $form.ShowDialog()
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
        PI_DA_Tuning_Parameters = $PIDATuningParameters
        PI_DA_ArchivesDisk = $PIDAArchivesDiskInfo
        PI_DA_Point_Sources = $PIDAPointSourcesInfo
        PI_AF_Informations = $PIAFInformations
    }

    # Gestion du rapport - console + fichier
    Show-Report -AuditReport $AuditReport
    Write-ReportToFile -AuditReport $AuditReport
    # Show-FormattedReportGUI
}

# APPEL FONCTION MAIN
Main