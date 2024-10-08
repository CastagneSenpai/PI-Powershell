# Créer une instance de l'application Outlook
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Accéder au calendrier principal dans Outlook
$mainCalendar = $Namespace.GetDefaultFolder(9) # 9 correspond au calendrier principal

# Récupérer les éléments du calendrier principal
$mainCalendarItems = $mainCalendar.Items
$mainCalendarItems.IncludeRecurrences = $true

# Créer une table de hachage pour stocker les événements uniques
$eventHash = @{}

# Parcourir tous les événements dans le calendrier
foreach ($item in $mainCalendarItems) {
    # Créer une clé unique basée sur le sujet et la date de début/fin (sans secondes)
    $startDate = $item.Start.ToString('yyyy-MM-dd HH:mm')
    $endDate = $item.End.ToString('yyyy-MM-dd HH:mm')
    $key = "$($item.Subject)|$startDate|$endDate"

    # Vérifier si la clé existe déjà dans la table de hachage
    if (-not $eventHash.ContainsKey($key)) {
        # Si ce n'est pas un doublon, on l'ajoute à la table de hachage
        $eventHash[$key] = $item
    } else {
        # Si c'est un doublon, on le supprime
        try {
            $item.Delete()
            Write-Host "Doublon supprimé : $($item.Subject) de $($item.Start) à $($item.End)"
        } catch {
            Write-Host "Erreur lors de la suppression de l'événement : $_"
        }
    }
}

Write-Host "Nettoyage des doublons terminé."
