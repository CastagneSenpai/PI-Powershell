param (
    [Parameter(Mandatory=$true)]
    [string]$username,
    [Parameter(Mandatory=$false)]
    [string]$basePath = "E:\01_ESPACE_UTILISATEURS",
    [Parameter(Mandatory=$false)]
    [string]$sourcesPath = "E:\00_ESPACE_ADMIN\02_GESTION DES UTILISATEURS\02_KIT_NOUVEL_UTILISATEUR"
)

Import-Module (Join-Path $PSScriptRoot 'lib\logs.psm1')
Clear-Host

function New-Repositories{
    param (
        $folders,
        $userPath
    )

    # CREATION DU REPERTOIRE RACINE DE L'UTILISATEUR
    try {
        $directory = New-Item -Path $userPath -ItemType Directory -ErrorAction Stop | Out-Null
        if ($directory -and $directory.PSIsContainer) {
            Write-Log -v_Message "R�pertoire utilisateur cr�� � l'emplacement $userPath" -v_ConsoleOutput -v_LogLevel "SUCCESS"
        }
    } catch {
        Write-Log -v_Message "Erreur lors de la cr�ation du r�pertoire utilisateur � l'emplacement $userPath : $_" -v_ConsoleOutput -v_LogLevel "ERROR"
    }
    
    # CREATION DES SOUS REPERTOIRES
    foreach ($folder in $folders) {
        try {
            $folderPath = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes((Join-Path -Path $userPath -ChildPath $folder)))
            New-Item -Path $folderPath -ItemType Directory -ErrorAction Stop | Out-Null
            Write-Log -v_Message "Dossier $folder cr�� � l'emplacement $folderPath" -v_ConsoleOutput -v_LogLevel "SUCCESS"
        } catch {
            Write-Log -v_Message "Erreur lors de la cr�ation du dossier $folder � l'emplacement $folderPath : $_" -v_ConsoleOutput -v_LogLevel "ERROR"
            exit
        }
    }
}

function Copy-Content{
    param (
        $userPath,
        $sourcesPath
    )

    try {
        # Copie les r�pertoires du r�pertoire source vers le r�pertoire utilisateur
        Get-ChildItem -Path $sourcesPath -Directory | ForEach-Object {
            Robocopy $_.FullName (Join-Path $userPath $_.Name) /E /Z /NFL /NDL /NP
        }
        Write-Log -v_Message "Sources copi�es dans $userPath." -v_ConsoleOutput -v_LogLevel "SUCCESS"
    }
    catch {
        Write-Log -v_Message "Erreur dans la copie des sources dans $userPath : $_" -v_ConsoleOutput -v_LogLevel "ERROR"
    }
}

function Update-Content {
    param (
        $userPath,
        $sourcesPath
    )
    try {
        Write-Log -v_Message "Mise � jour du contenu dans $userPath." -v_ConsoleOutput -v_LogLevel "INFO"
        Robocopy $sourcesPath $userPath /E /XC /XN /XO /Z /NFL /NDL /NP  # /XC /XN /XO pour ne copier que les fichiers plus r�cents ou non existants
        Write-Log -v_Message "Mise � jour termin�e avec succ�s dans $userPath." -v_ConsoleOutput -v_LogLevel "SUCCESS"
    } catch {
        Write-Log -v_Message "Erreur lors de la mise � jour dans $userPath : $_" -v_ConsoleOutput -v_LogLevel "ERROR"
    }
}

function Main {
    $userPath = Join-Path -Path $basePath -ChildPath $username

    if (Test-Path $userPath) {
        Write-Log -v_Message "L'utilisateur $username existe d�j�." -v_ConsoleOutput -v_LogLevel "WARN"
        $response = Read-Host "Souhaitez-vous mettre � jour l'arborescence de l'utilisateur ? (Y/N)"
        if ($response -eq "Y") {
            Update-Content -userPath $userPath -sourcesPath $sourcesPath
        } else {
            Write-Log -v_Message "Aucune action n'a �t� entreprise." -v_ConsoleOutput -v_LogLevel "INFO"
        }
    } else {
        $foldersToCreate = Get-ChildItem -Path $sourcesPath -Directory | Select-Object -ExpandProperty Name
        New-Repositories -folders $foldersToCreate -userPath $userPath
        Copy-Content -sourcesPath $sourcesPath -userPath $userPath
    }
}

Main -username $username
