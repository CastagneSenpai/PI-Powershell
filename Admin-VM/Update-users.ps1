param (
    [Parameter(Mandatory=$false)]
    [string]$basePath = "E:\01_ESPACE_UTILISATEURS",
    [Parameter(Mandatory=$false)]
    [string]$sourcesPath = "E:\00_ESPACE_ADMIN\02_GESTION DES UTILISATEURS\02_KIT_NOUVEL_UTILISATEUR"
)

Import-Module (Join-Path $PSScriptRoot 'lib\logs.psm1')
Clear-Host

function Update-UserRepositories {
    param (
        $userPath,
        $sourcesPath
    )

    Write-Log -v_Message "Mise à jour de $userPath à partir de $sourcesPath." -v_ConsoleOutput -v_LogLevel "INFO"
    try {
        # Utilisation de Robocopy pour mettre à jour les fichiers
        Robocopy $sourcesPath $userPath /E /XC /XN /XO /Z /NFL /NDL /NP  # /XC /XN /XO pour ne copier que les fichiers plus récents ou non existants
        Write-Log -v_Message "Mise à jour terminée avec succès dans $userPath." -v_ConsoleOutput -v_LogLevel "SUCCESS"
    } catch {
        Write-Log -v_Message "Erreur lors de la mise à jour de $userPath : $_" -v_ConsoleOutput -v_LogLevel "ERROR"
    }
}

function Main {
    # Obtenir tous les répertoires utilisateurs dans le chemin de base
    $userDirs = Get-ChildItem -Path $basePath -Directory

    foreach ($userDir in $userDirs) {
        $userPath = $userDir.FullName
        Write-Log -v_Message "Traitement de l'utilisateur $($userDir.Name)" -v_ConsoleOutput -v_LogLevel "INFO"

        if (Test-Path $userPath) {
            Update-UserRepositories -userPath $userPath -sourcesPath $sourcesPath
        } else {
            Write-Log -v_Message "Le chemin $userPath n'existe pas, aucun traitement effectué." -v_ConsoleOutput -v_LogLevel "WARN"
        }
    }
}

Main
