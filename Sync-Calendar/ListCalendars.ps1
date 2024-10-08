# Créer une instance de l'application Outlook
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Fonction pour lister les dossiers récursivement
function List-Folders {
    param (
        [Microsoft.Office.Interop.Outlook.MAPIFolder]$Folder
    )
    Write-Host "Dossier: $($Folder.Name)"
    foreach ($subFolder in $Folder.Folders) {
        List-Folders $subFolder
    }
}

# Lister tous les dossiers dans le namespace MAPI
foreach ($folder in $Namespace.Folders) {
    List-Folders $folder
}
