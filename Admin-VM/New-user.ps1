param (
    [Parameter(Mandatory=$true)]
    [string]$username,
    [Parameter(Mandatory=$false)]
    [string]$basePath = "E:\01_ESPACE_UTILISATEURS",
    [Parameter(Mandatory=$false)]
    [string]$sourcesPath = "E:\00_ESPACE_ADMIN\02_GESTION DES UTILISATEURS\02_KIT_NOUVEL_UTILISATEUR"
)

Import-Module (Join-Path $PSScriptRoot '..\lib\logs.psm1')
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
            Write-Log -v_Message "Répertoire utilisateur créé à l'emplacement $userPath" -v_ConsoleOutput -v_LogLevel "SUCCESS"
        }
    } catch {
        Write-Log -v_Message "Erreur lors de la création du répertoire utilisateur à l'emplacement $userPath : $_" -v_ConsoleOutput -v_LogLevel "ERROR"
    }
    
    # CREATION DES SOUS REPERTOIRES
    foreach ($folder in $folders) {
        try {
            $folderPath = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes((Join-Path -Path $userPath -ChildPath $folder)))
            New-Item -Path $folderPath -ItemType Directory -ErrorAction Stop | Out-Null
            Write-Log -v_Message "Dossier $folder cree a l'emplacement $folderPath" -v_ConsoleOutput -v_LogLevel "SUCCESS"
        } catch {
            Write-Log -v_Message "Erreur lors de la création du dossier $folder à l'emplacement $folderPath : $_" -v_ConsoleOutput -v_LogLevel "ERROR"
            exit
        }
    }
     
}

function Copy-Content{
    param (
        $userPath,
        $sourcesPath
    )

    $folders = Get-ChildItem -path $userPath

    try{
        Robocopy (Join-Path $sourcesPath "Formations") $folders[0].FullName /MIR /Z /NFL /NDL /NP
        Robocopy (Join-Path $sourcesPath "Documents") $folders[1].FullName /MIR /Z /NFL /NDL /NP
        Robocopy (Join-Path $sourcesPath "Projets") $folders[2].FullName /MIR /Z /NFL /NDL /NP
        Write-Log -v_Message "Sources copiées dans $userPath."  -v_ConsoleOutput -v_LogLevel "SUCCESS"
    }
    catch{
        Write-Log -v_Message "Erreur dans la copie des sources dans $userPath."  -v_ConsoleOutput -v_LogLevel "ERROR"
    }
    
}

function Update-Content {
    param (
        $userPath,
        $sourcesPath
    )
    try {
        Write-Log -v_Message "Mise à jour du contenu dans $userPath." -v_ConsoleOutput -v_LogLevel "INFO"
        Robocopy $sourcesPath $userPath /MIR /Z /NFL /NDL /NP /XO  # /XO permet de ne copier que les fichiers plus récents ou non existants
        Write-Log -v_Message "Mise à jour terminée avec succès dans $userPath." -v_ConsoleOutput -v_LogLevel "SUCCESS"
    } catch {
        Write-Log -v_Message "Erreur lors de la mise à jour dans $userPath : $_" -v_ConsoleOutput -v_LogLevel "ERROR"
    }
}

function Main {
    $userPath = Join-Path -Path $basePath -ChildPath $username
    $foldersToCreate = @("01_Espace_Formation", "02_Espace_Documents", "03_Espace_Projets", "04_Espace_Personnel")

    if (Test-Path $userPath) {
        Write-Log -v_Message "L'utilisateur $username existe déjà." -v_ConsoleOutput -v_LogLevel "WARNING"
        $response = Read-Host "Souhaitez-vous mettre à jour l'arborescence de l'utilisateur ? (Y/N)"
        if ($response -eq "Y") {
            Update-Content -userPath $userPath -sourcesPath $sourcesPath
        } else {
            Write-Log -v_Message "Aucune action n'a été entreprise." -v_ConsoleOutput -v_LogLevel "INFO"
        }
    } else {
        New-Repositories -folders $foldersToCreate -userPath $userPath
        Copy-Content -sourcesPath $sourcesPath -userPath $userPath
    }
}

Main -username $username
