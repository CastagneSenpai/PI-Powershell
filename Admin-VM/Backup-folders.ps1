<#
.SYNOPSIS
    Script de backup des répertoires utilisateurs avec gestion des tâches planifiées.

.DESCRIPTION
    Ce script PowerShell réalise une sauvegarde des répertoires utilisateurs en les compressant au format ZIP, puis nettoie les anciennes sauvegardes. Il vérifie également la présence d'une tâche planifiée pour exécuter le script régulièrement et la crée si nécessaire.

.NOTES
    Auteur     : Romain Castagné
    Société    : CGI
    Date       : 22 Août 2024
    Version    : 1.0

.PARAMETER SourceDir
    Répertoire source contenant les dossiers utilisateurs à sauvegarder.

.PARAMETER BackupDir
    Répertoire de destination pour les fichiers ZIP de sauvegarde.

.PARAMETER DaysToKeep
    Nombre de jours pendant lesquels les fichiers ZIP de sauvegarde doivent être conservés.

.EXAMPLE
    .\Backup-folders.ps1 -SourceDir "E:\01_ESPACE_UTILISATEURS" -BackupDir "E:\00_ESPACE_ADMIN\02_GESTION DES UTILISATEURS\03_BACKUP_DES_ESPACES" -DaysToKeep 5

    Lance le script pour sauvegarder les répertoires utilisateurs, en conservant les sauvegardes pour une durée de 5 jours.
#>

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Import de la fonction Write-Log
Import-Module "$PSScriptRoot\lib\logs.psm1"

# Fonction pour zipper un répertoire
Function Compress-Folder {
    param (
        [string]$FolderPath,
        [string]$DestinationPath
    )

    try {
        $zipFileName = (Join-Path -Path $DestinationPath -ChildPath ("$([System.IO.Path]::GetFileNameWithoutExtension($FolderPath))_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".zip"))
        Write-Log -v_Message "Zipping folder $FolderPath to $zipFileName" -v_LogLevel "INFO" -v_ConsoleOutput
        Compress-Archive -Path $FolderPath -DestinationPath $zipFileName -Force
    }
    catch {
        Write-Log -v_Message "Failed to zip folder $FolderPath : $_" -v_LogLevel "ERROR" -v_ConsoleOutput
    }

    return $zipFileName
}

# Fonction pour nettoyer les anciens fichiers
Function Remove-OldBackups {
    param (
        [string]$BackupDir,
        [int]$DaysToKeep = 5
    )

    try {
        $cutoffDate = (Get-Date).AddDays(-$DaysToKeep)
        $filesToDelete = Get-ChildItem -Path $BackupDir -Filter *.zip | Where-Object { $_.LastWriteTime -lt $cutoffDate }
        foreach ($file in $filesToDelete) {
            Write-Log -v_Message "Deleting old backup file $($file.FullName)" -v_LogLevel "INFO" -v_ConsoleOutput
            Remove-Item $file.FullName -Force
        }
    }
    catch {
        Write-Log -v_Message "Failed to clean up old backups: $_" -v_LogLevel "ERROR" -v_ConsoleOutput
    }
}

# Fonction pour créer ou vérifier la tâche planifiée
Function Install-ScheduledTask {
    param (
        [string]$TaskName,
        [string]$ScriptPath
    )

    $taskExists = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

    if (!$taskExists) {
        try {
            Write-Log -v_Message "Creating scheduled task $TaskName" -v_LogLevel "INFO" -v_ConsoleOutput
            $action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$ScriptPath`""
            $trigger = New-ScheduledTaskTrigger -Daily -At 2am
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to create scheduled task $TaskName : $_"  
            exit
        }
    }
    else {
        Write-Log -v_Message "Scheduled task $TaskName already exists. No action needed." -v_LogLevel "INFO" -v_ConsoleOutput
    }
}

Function Main {
    param (
        [string]$SourceDir = "E:\01_ESPACE_UTILISATEURS",
        [string]$BackupDir = "E:\00_ESPACE_ADMIN\02_GESTION DES UTILISATEURS\03_BACKUP_DES_ESPACES",
        [int]$DaysToKeep = 5
    )

    # Vérification et création de la tâche planifiée si nécessaire
    Install-ScheduledTask -TaskName "BACKUP_USERS_FOLDERS" -ScriptPath (join-path $PSScriptRoot "Backup-folders.ps1") -ErrorAction Stop
    
    
    # Parcours de chaque dossier enfant dans le répertoire source
    Get-ChildItem -Path $SourceDir -Directory | ForEach-Object {
        $userDir = $_.FullName
        Write-Log -v_Message "Processing directory $userDir" -v_LogLevel "INFO" -v_ConsoleOutput

        # Création d'un fichier zip du répertoire utilisateur
        Compress-Folder -FolderPath $userDir -DestinationPath $BackupDir

        # Nettoyage des anciens fichiers dans le répertoire de backup
        Remove-OldBackups -BackupDir $BackupDir -DaysToKeep $DaysToKeep
    }

    Write-Log -v_Message "Backup process completed successfully." -v_LogLevel "SUCCESS" -v_ConsoleOutput
}

# Appel de la fonction principale
Main
