<#
.SYNOPSIS
    Script PowerShell pour synchroniser les événements du calendrier Google avec le calendrier principal Outlook, en ajoutant et supprimant les réunions si nécessaire.

.DESCRIPTION
    Ce script récupère les événements du calendrier spécifié dans "Internet Calendars" et les ajoute au calendrier principal Outlook.
    Il vérifie si les événements existent déjà pour éviter les doublons et supprime ceux qui ne sont plus présents dans le calendrier Google.

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
    
    # Créer une instance Outlook
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    
    # Accéder au dossier "Internet Calendars"
    $internetCalendars = $Namespace.Folders.Item("Internet Calendars")
    if ($null -eq $internetCalendars) {
        Add-Content -path $LogPath -value "Le dossier 'Internet Calendars' n'a pas été trouvé."
        exit
    }
    
    # Accéder au calendrier Google
    $calendar = $internetCalendars.Folders.Item($CalendarEmail)
    if ($null -eq $calendar) {
        Add-Content -path $LogPath -value "Le calendrier spécifié n'a pas été trouvé."
        exit
    }
    
    # Récupérer les événements du calendrier Google
    $items = $calendar.Items | Where-Object { $_.Start -gt (Get-Date).AddDays(-1)} | Sort-Object Start
    
    # Accéder au calendrier principal Outlook
    $mainCalendar = $Namespace.GetDefaultFolder(9)
    $mainCalendarItems = $mainCalendar.Items | Sort-Object Start
    
    $googleEvents = @{}
    
    # Synchroniser les réunions
    foreach ($item in $items) {
        $eventKey = "$($item.Subject)-$($item.Start)-$($item.End)"
        $googleEvents[$eventKey] = $true
        
        $eventExists = $false
        foreach ($mainItem in $mainCalendarItems) {
            if ($mainItem.Subject -eq "SYENSQO meeting" -and $mainItem.Start -eq $item.Start -and [math]::Abs(($mainItem.End - $item.End).TotalMinutes) -le 5) {
                $eventExists = $true
                break
            }
        }
        
        if (-not $eventExists) {
            $newMeeting = $Outlook.CreateItem(1)
            $newMeeting.Subject = if ($item.Subject -eq "Busy") { "SYENSQO meeting" } else { $item.Subject }
            $newMeeting.Start = $item.Start
            $newMeeting.End = $item.End
            $newMeeting.Body = "Ajouté depuis le calendrier Internet"
            $newMeeting.ReminderSet = $true
            $newMeeting.ReminderMinutesBeforeStart = 1
            
            try {
                $newMeeting.Save()
                Add-Content -path $LogPath -value "Réunion créée : $($newMeeting.Subject) de $($newMeeting.Start) à $($newMeeting.End)"
            } catch {
                Add-Content -path $LogPath -value "Erreur lors de la création de la réunion : $_"
            }
        }
    }
    
    # Suppression des réunions obsolètes
    foreach ($mainItem in $mainCalendarItems) {
        if ($mainItem.Subject -eq "SYENSQO meeting") {
            $foundMatch = $false
            foreach ($googleItem in $items) {
                if ($mainItem.Start -eq $googleItem.Start -and [math]::Abs(($mainItem.End - $googleItem.End).TotalMinutes) -le 25) {
                    $foundMatch = $true
                    break
                }
            }
            if (-not $foundMatch) {
                try {
                    $mainItem.Delete()
                    Add-Content -path $LogPath -value "Réunion supprimée : $($mainItem.Subject) de $($mainItem.Start) à $($mainItem.End)"
                } catch {
                    Add-Content -path $LogPath -value "Erreur lors de la suppression de la réunion : $_"
                }
            }
        }
    }
}
catch {
    Add-Content -path $LogPath -value "[$(Get-Date)] Erreur ligne $($_.InvocationInfo.ScriptLineNumber) :: $($_.Exception.Message)"
}
