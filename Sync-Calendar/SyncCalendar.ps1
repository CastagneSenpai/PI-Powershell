# Fonction pour verifier si une tâche planifiee existe
function Check-ScheduledTaskExists {
    param (
        [string]$TaskName
    )
    
    try {
        $task = Get-ScheduledTask | Where-Object { $_.TaskName -eq $TaskName }
        return $task -ne $null
    } catch {
        Write-Host "Erreur lors de la verification de la tâche planifiee : $_"
        return $false
    }
}

# Vérifie si le script est exécuté avec des droits administratifs
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # Relance le script avec des droits administratifs
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs
    exit
}

# Si la tâche planifiee n'existe pas, la creer
$taskName = "Update SYENSQO Calendar"

if (-not (Check-ScheduledTaskExists $taskName)) {
    Write-Host "La tâche planifiee '$taskName' n'existe pas, creation en cours..."

    # Creer une nouvelle tâche planifiee avec un declencheur au demarrage et une repetition
    $action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File 'C:\Scripts\SyncCalendar.ps1'"
    $trigger = New-ScheduledTaskTrigger -AtStartup

    $principal = New-ScheduledTaskPrincipal -UserId "S-1-5-21-3641078771-3653456904-245653651-1325287" -LogonType Interactive -RunLevel Highest

    # Enregistrer la tâche
    Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName $taskName -Description "Mise a jour du calendrier SYENSQO"

    Write-Host "Tâche planifiee '$taskName' creee avec succes."
} else {
    Write-Host "La tâche planifiee '$taskName' existe deja."
}

# ------------------- Suite du script existant -------------------

# Creer une instance de l'application Outlook
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Acceder au dossier "Internet Calendars"
$internetCalendars = $Namespace.Folders.Item("Internet Calendars")

# Verifier si le dossier est bien recupere
if ($internetCalendars -eq $null) {
    Write-Host "Le dossier 'Internet Calendars' n'a pas ete trouve."
    exit
}

# Acceder au calendrier dans le dossier "Internet Calendars"
$calendar = $internetCalendars.Folders.Item("romain.castagne-ext@syensqo.com") # Nom du calendrier

# Verifier si le calendrier est bien recupere
if ($calendar -eq $null) {
    Write-Host "Le calendrier specifie dans 'Internet Calendars' n'a pas ete trouve."
    exit
}

# Recuperer tous les elements du calendrier, sans filtrer par date
$items = $calendar.Items
$items.IncludeRecurrences = $true

# Trier les elements par date de debut
$items.Sort("[Start]")

# Acceder au calendrier principal dans Outlook pour ajouter des evenements
$mainCalendar = $Namespace.GetDefaultFolder(9) # 9 correspond au calendrier principal

# Recuperer les elements du calendrier principal
$mainCalendarItems = $mainCalendar.Items
$mainCalendarItems.IncludeRecurrences = $true

# Trier les elements du calendrier principal par date de debut
$mainCalendarItems.Sort("[Start]")

# Creer des evenements dans le calendrier principal a partir des elements recuperes
foreach ($item in $items) {
    # Verifier si l'evenement existe deja dans le calendrier principal
    $eventExists = $false
    foreach ($mainItem in $mainCalendarItems) {
        # Comparer le sujet, la date de debut et la date de fin
        if ($mainItem.Subject -eq "SYENSQO meeting" -and $mainItem.Start -eq $item.Start -and $mainItem.End -eq $item.End) {
            $eventExists = $true
            break
        }
    }

    # Si l'evenement n'existe pas, on le cree
    if (-not $eventExists) {
        # Creer une reunion
        $newMeeting = $mainCalendarItems.Add(1) # 1 pour un rendez-vous

        # Renommer l'evenement si le sujet est "Busy"
        if ($item.Subject -eq "Busy") {
            $newMeeting.Subject = "SYENSQO meeting"
        } else {
            $newMeeting.Subject = $item.Subject
        }

        # Assigner les dates de debut et de fin
        $newMeeting.Start = $item.Start
        $newMeeting.End = $item.End
        
        $newMeeting.Body = "Ajoute depuis le calendrier Internet"
        $newMeeting.ReminderSet = $true
        $newMeeting.ReminderMinutesBeforeStart = 15 # Rappel 15 minutes avant

        # Sauvegarder la reunion
        try {
            $newMeeting.Save()
            Write-Host "Reunion creee : $($newMeeting.Subject) de $($newMeeting.Start) a $($newMeeting.End)"
        } catch {
            Write-Host "Erreur lors de la creation de la reunion : $_"
        }
    } else {
        Write-Host "La reunion du $($mainItem.Start) au $($mainItem.End) existe deja."
    }
}
