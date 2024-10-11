<#
.SYNOPSIS
    Script PowerShell pour synchroniser les événements du calendrier romain.castagne-ext@syensqo.com avec le calendrier principal Outlook.

.DESCRIPTION
    Ce script récupère les événements du calendrier spécifié dans "Internet Calendars" et les ajoute au calendrier principal Outlook.
    Il vérifie si les événements existent déjà pour éviter les doublons et enregistre les événements créés dans un fichier log.

.PARAMETER CalendarEmail
    Adresse e-mail du calendrier Internet à récupérer.

.PARAMETER LogPath
    Chemin du fichier log où les événements créés et les erreurs seront enregistrés.

.EXAMPLE
    .\SyncCalendar.ps1 -CalendarEmail "romain.castagne-ext@syensqo.com" -LogPath "C:\Scripts\log.txt"
    Exécute le script en synchronisant le calendrier spécifié et en enregistrant les logs dans le fichier spécifié.

.NOTES
    Auteur: Romain Castagné
    Version: 1.0
    Date: $(Get-Date -Format "yyyy-MM-dd")
    Tâche planifiée: Update SYENSQO Calendar.xml doit être importée dans le planificateur de tâches avec l'adresse de calendrier Outlook appropriée.
#>

param (
    [string]$CalendarEmail = "romain.castagne-ext@syensqo.com",
    [string]$LogPath = (Join-Path $PSScriptRoot "log.txt")
)

try {
    
    # Démarrer Outlook
    while (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
        Add-Content -path $LogPath -value "Ouverture de Outlook en cours ..."
        Start-Sleep -Seconds 2
    }
    
    # Créer une instance de l'application Outlook
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    
    # Accéder au dossier "Internet Calendars"
    $internetCalendars = $Namespace.Folders.Item("Internet Calendars")
    
    # Vérifier si le dossier est bien récupéré
    if ($null -eq $internetCalendars) {
        Add-Content -path $LogPath -value "Le dossier 'Internet Calendars' n'a pas été trouvé."
        exit
    }
    
    # Accéder au calendrier dans le dossier "Internet Calendars"
    $calendar = $internetCalendars.Folders.Item($CalendarEmail) # Utiliser le paramètre CalendarEmail
    
    # Vérifier si le calendrier est bien récupéré
    if ($null -eq $calendar) {
        Add-Content -path $LogPath -value "Le calendrier spécifié dans 'Internet Calendars' n'a pas été trouvé."
        exit
    }
    
    # Récupérer tous les éléments du calendrier, sans filtrer par date
    $items = $calendar.Items | Where-Object { $_.Start -gt (Get-Date).AddDays(-1)} | Sort-Object Start
    
    # Accéder au calendrier principal dans Outlook pour ajouter des événements
    $mainCalendar = $Namespace.GetDefaultFolder(9) # 9 correspond au calendrier principal
    
    # Récupérer les éléments du calendrier principal
    $mainCalendarItems = $mainCalendar.Items | Sort-Object Start
    
    # Créer des événements dans le calendrier principal à partir des éléments récupérés
    foreach ($item in $items) {
        # Vérifier si l'événement existe déjà dans le calendrier principal
        $eventExists = $false
        foreach ($mainItem in $mainCalendarItems) {
            if ($mainItem.Subject -eq "SYENSQO meeting" -and $mainItem.Start -eq $item.Start -and $mainItem.End -eq $item.End) {
                $eventExists = $true
                break
            }
        }
    
        # Si l'événement n'existe pas, on le crée
        if (-not $eventExists) {
            # Créer une réunion
            $newMeeting = $Outlook.CreateItem(1) # 1 pour un rendez-vous
    
            # Renommer l'événement si le sujet est "Busy"
            if ($item.Subject -eq "Busy") {
                $newMeeting.Subject = "SYENSQO meeting"
            } else {
                $newMeeting.Subject = $item.Subject
            }
    
            # Assigner les dates de début et de fin
            $newMeeting.Start = $item.Start
            $newMeeting.End = $item.End
            
            $newMeeting.Body = "Ajouté depuis le calendrier Internet"
            $newMeeting.ReminderSet = $true
            $newMeeting.ReminderMinutesBeforeStart = 15 # Rappel 15 minutes avant
    
            # Sauvegarder la réunion
            try {
                $newMeeting.Save()
                Add-Content -path $LogPath -value "Réunion créée : $($newMeeting.Subject) de $($newMeeting.Start) à $($newMeeting.End)"
            } catch {
                Add-Content -path $LogPath -value "Erreur lors de la création de la réunion : $_"
            }
        } else {
            Add-Content -path $LogPath -value "La réunion du $($mainItem.Start) au $($mainItem.End) existe déjà."
        }
    }
}
catch {
    Add-Content -path $LogPath -value "[$(Get-Date)] Erreur ligne $($_.InvocationInfo.ScriptLineNumber) :: $($_.Exception.Message)"
}