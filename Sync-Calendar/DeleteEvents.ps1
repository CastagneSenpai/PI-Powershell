# Creer une instance de l'application Outlook
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Acceder au calendrier principal dans Outlook
$mainCalendar = $Namespace.GetDefaultFolder(9) # 9 correspond au calendrier principal

# Recuperer les elements du calendrier principal
$mainCalendarItems = $mainCalendar.Items
$mainCalendarItems.IncludeRecurrences = $true

# Liste pour stocker les evenements a supprimer
$eventsToDelete = @()

# Trouver tous les evenements avec le sujet "SYENSQO meeting"
foreach ($item in $mainCalendarItems) {
    if ($item.Subject -eq "SYENSQO meeting") {
        $eventsToDelete += $item
    }
}

# Supprimer les evenements trouves
foreach ($event in $eventsToDelete) {
    try {
        Write-Host "evenement supprime : $($event.Subject) de $($event.Start)"
        $event.Delete()
    } catch {
        Write-Host "Erreur lors de la suppression de l'evenement : $_"
    }
}

if ($eventsToDelete.Count -eq 0) {
    Write-Host "Aucun evenement 'SYENSQO meeting' trouve a supprimer."
}
