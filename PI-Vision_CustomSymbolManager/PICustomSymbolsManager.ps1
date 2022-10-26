<#
.SYNOPSIS
Application script use to compare identify the custom symbols in a server, his validity, his version and extra files.

.DESCRIPTION
This script pulls all custom symbol files hashes for given PI Vision server list. 
It compare this informations to a referential file to identify the CS installed, their validity and their version. 
It generate a csv file as an output which contain the PI Vision server, the CS names identified, good/bad validity, good/bad version, and an error message if needed.
Naming pattern is: yyyy-MM-ddTHH-mm-ss_PICustomSymbolsAnalysis.csv

.PARAMETER inputFile_ServersList
Path to CSV file containing at least two columns "ServerName" (PI Vision server) and "PIVisionExtPath" (PI Vision server 'ext' folder path UNC format).

.PARAMETER inputFile_PICustomSymbolReferential
Path to CSV file containing at least four columns "Algorithm" (HashMode), "Hash" (checksum of the file), "Path" (Path from ext folder) and symbol (link a file to a referenced symbol).

.PARAMETER outputFolder
[OPTIONAL] Folder for output CSV file. To be set for custom target folder.
[DEFAULT] PSScriptRoot\output

.PARAMETER logFile
[OPTIONAL] Path to log file. To be set for custom target file.
[DEFAULT] PSScriptRoot\log\yyyy-MM-dd_PICustomSymbolManager.log

.NOTES
    FILE:    PICustomSymbolManager.ps1
    AUTHOR:  Romain CASTAGN� <romain.castagne@external.totalenergies.com>
    VERSION: 1.0

#> 

##############################################################################
# Main function with loop logic over several servers
##############################################################################

Function Main() {
    [CmdletBinding()]
    param (
    [Parameter(Mandatory=$false)][string]$inputFile_ServersList = (Join-Path $PSScriptRoot "input\Server_Path_List.csv"), #File with PI Vision servers list & associated path to ext/ folder
    [Parameter(Mandatory=$false)][string]$inputFile_PICustomSymbolReferential = (Join-Path $PSScriptRoot "input\Referential_PICustomSymbolFiles.csv"), #File with list of files & hash for each symbols
    [Parameter(Mandatory=$false)][string]$outputFolder = (Join-Path $PSScriptRoot "output\"), # Custom output folder if needed.
    [Parameter(Mandatory=$false)][string]$logFile = (Join-Path $PSScriptRoot "logs\$(Get-Date -Format 'yyyy-MM-ddTHH-mm-ss')_PICustomSymbolManager.log") # Custom log file if needed.
    )

    # LOGS     
    Clear-Host
    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Entering PI Custom Symbols Manager script."
    Write-Log -v_LogFile $logFile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "inputFile_ServersList = $inputFile_ServersList"
    Write-Log -v_LogFile $logFile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "inputFile_PICustomSymbolReferential = $inputFile_PICustomSymbolReferential"
    Write-Log -v_LogFile $logFile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "outputFolder = $outputFolder"
    Write-Log -v_LogFile $logFile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "logFile = $logFile"
    
    # VERIFYING THAT WE HAVE THE TWO REQUIRED INPUT FILE AVAILABLE.
    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Check the availability of the input files."
    if (!(Test-Path $inputFile_ServersList)) {
        Write-Log -v_LogFile $logFile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "An input file is missing : $inputFile_ServersList."
        return 
    }
    if (!(Test-Path $inputFile_PICustomSymbolReferential)) {
        Write-Log -v_LogFile $logFile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "An input file is missing : $inputFile_PICustomSymbolReferential."
        return 
    }

    # MAKING OUTPUT FOLDER IF NOT EXISTS
    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Manage output folder and file."
    if (!(Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder
    }
    # CREATE AN OUTPUT FILE IN OUTPUT SPECIFIED DIRECTORY WHICH WILL CONTAIN THE RESULT OF THE PI CUSTOM SYMBOLS MANAGER
    [string]$outputFile = (Join-Path $outputFolder ("$(Get-Date -Format 'yyyy-MM-dd')_PICustomSymbolsAnalysis.csv"))

    # CREATE A CUSTOM OBJECT WHICH IS THE STRUCTURE OF THE OUTPUT ANALYSIS
    $Report = @()
    
    # LOOPING OVER INPUT FILE ROWS
    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Import the server list from input file."
    $PIVisionServers = Import-Csv $inputFile_ServersList

    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Import the referencial hash codes file from input file."
    $ReferentialFiles = Import-Csv $inputFile_PICustomSymbolReferential -Delimiter ";"
    # Write-Log -v_LogFile $logFile -v_LogLevel WARN -v_ConsoleOutput -v_Message "ReferentialFiles = $($ReferentialFiles.Algorithm.Fist), $($ReferentialFiles.Hash[0]), $($ReferentialFiles.Path[0]), $($ReferentialFiles.symbol[0])."
    
    
    foreach ($PIVisionServer in $PIVisionServers) 
    {
        Write-EmptyLine -v_LogFile $logFile
        Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Start processing server $($PIVisionServer.ServerName)."

        # GET REMOTE FILES AND GET THEIR HASH CODE.
        Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Get remote files and hash codes."
        $LocalCSFiles = Get-RemoteFilesHash $PIVisionServer.ServerName $PIVisionServer.PIVisionExtPath

        Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Identify which symbols are present in the server."
        foreach($File in $LocalCSFiles)
        {
            # FOR EACH SYMBOLS 
            if($File.Path -like "*.js" -And $File.Path -notlike "*libraries*")
            {
                # CLEAR VARIABLES USED FOR LAST SYMBOL
                if( Get-Variable -Name Is* ) { Clear-Variable Is* }
                if( Get-Variable -Name ErrorMsg* ) { Clear-Variable ErrorMsg }


                # GET THE SYMBOL NAME
                $SymbolNameIndex = $File.Path.IndexOf("sym-") + 4  # Do not take the "sym-"
                $SymbolName = $File.Path.Substring($SymbolNameIndex, $File.Path.Length - $SymbolNameIndex -3)
                Write-EmptyLine -v_LogFile $logFile
                Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$($PIVisionServer.ServerName) : Symbol found = $SymbolName"
                
                # DETERMINES WHETHER THE SYMBOL IS REFERENCED IN THE REFERENCIAL INPUT FILE
                $ReferentialFilesFiltered = $ReferentialFiles | Where{ $_.Symbol -like "$SymbolName" }
                if([string]::IsNullOrEmpty(($ReferentialFilesFiltered.Symbol | Select-Object -First 1)))
                {
                    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$($PIVisionServer.ServerName) : $SymbolName is NOT referenced in the referential file."
                    $IsReferencedSymbol = $false
                    $IsValidFiles = $false
                    $IsValidVersion = $false
                    $ErrorMsg = "The symbol $SymbolName is not referenced in the referencial file."
                }
                else
                {
                    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$($PIVisionServer.ServerName) : $SymbolName is referenced."
                    $IsReferencedSymbol = $true

                    [string]$remoteRootPath = (Join-Path "\\$($PIVisionServer.ServerName)\" $PIVisionServer.PIVisionExtPath.Replace(':','$'))

                    # DETERMINES WHETHER ALL THE FILES OF THE SYMBOL ARE PRESENT - IS VALID SYMBOL                    
                    foreach($CurrentReferentialFile in $ReferentialFilesFiltered)
                    {
                        if ($LocalCSFiles.Path -notcontains (Join-path $remoteRootPath $CurrentReferentialFile.Path))
                        {
                            Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$($PIVisionServer.ServerName) : Missing files for $SymbolName symbol."
                            $IsValidFiles = $false
                            $IsValidVersion = $false
                            $ErrorMsg = "File $($CurrentReferentialFile.Path) not found."
                            
                            # GO TO NEXT SYMBOL
                            break
                        }
                        $IsValidFiles = $true
                    }

                    if($IsValidFiles)
                    {
                        Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$($PIVisionServer.ServerName) : All the files for $SymbolName symbol are present."

                        # DETERMINES WHETHER ALL THE HASH FILES ARE THE SAME THAN THE REFERENCIAL - IS VALID VERSION OF THE SYMBOL
                        foreach($CurrentReferentialFile in $ReferentialFilesFiltered)
                        {
                            $CurrentFileHashInServer = Get-FileHash (Join-path $remoteRootPath $CurrentReferentialFile.Path) -Algorithm SHA512
                            if($CurrentFileHashInServer.Hash -ne $CurrentReferentialFile.Hash)
                            {
                                Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$($PIVisionServer.ServerName) : Version of file $($CurrentReferentialFile.Path) is not the same in the referencial for the symbol $SymbolName."
                                $IsValidVersion = $false
                                $ErrorMsg = "Version of file $($CurrentReferentialFile.Path) is not the same in the referencial."
                                
                                # GO TO NEXT SYMBOL
                                break
                            }
                            $IsValidVersion = $true
                            $ErrorMsg = "Symbol OK"
                        }
                        
                        if($IsValidVersion -eq $true) { 
                            Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$($PIVisionServer.ServerName) : All the files for $SymbolName symbol are in the good version."
                        } 
                    }
                }
                
                # ADD TO THE REPORT CUSTOM OBJECT THE CURRENT SYMBOL ANALYSIS RESULT
                Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$($PIVisionServer.ServerName) : Add the symbol $SymbolName to the report file."         
                $Report += [PSCustomObject]@{
                    PIVisionServer = $PIVisionServer.ServerName; 
                    CustomSymbolPath = $PIVisionServer.PIVisionExtPath; 
                    SymbolName = $SymbolName; 
                    IsReferencedSymbol = $IsReferencedSymbol;
                    IsValidFiles = "$IsValidFiles";
                    IsValidVersion = "$IsValidVersion"
                    ErrorMsg = "$ErrorMsg";
                }
            }
        }
        Write-Log -v_LogFile $logFile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "Stop processing server $($PIVisionServer.ServerName)."
    }   
    
    # Export the report into a CSV file
    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Export the report into a CSV file."
    $Report | Export-Csv -UseCulture -Encoding UTF8 -NoTypeInformation $outputFile
	$Report | ConvertTo-Html -Head "PI Custom Symbol Manager - Status of CS versions" | Out-File (Join-Path $outputFolder ("$(Get-Date -Format 'yyyy-MM-dd')_PICustomSymbolsAnalysis.html"))
	
    # End of PI Custom Symbol Manager analysis
    Write-Log -v_LogFile $logFile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Exiting PI Custom Symbols Manager script."

    Read-Host "Exiting PI Custom Symbol Manager..."
}

##############################################################################
# Function to get Hash Files for a single server
##############################################################################
Function Get-RemoteFilesHash ([string]$serverName, [string]$localRootPath) {
    Write-Log -v_LogFile $logFile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "Entering Get-RemoteFilesHash function."
    Write-Log -v_LogFile $logFile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "serverName='$serverName'"
    Write-Log -v_LogFile $logFile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "localRootPath='$localRootPath'"

    [string]$remoteRootPath = (Join-Path "\\$serverName\" $localRootPath.Replace(':','$'))
    if (!(Test-Path $remoteRootPath)) {
        Write-Log "ERROR" "Not able to list files in '$remoteRootPath'."
        return
    }
    # We retrieve all files recursively (subfolders) and get their SHA512 hash
    $remoteFiles = Get-ChildItem $remoteRootPath -Recurse | % { Get-FileHash $_.FullName -Algorithm SHA512 }

    Write-Log -v_LogFile $logFile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "Exiting Get-RemoteFilesHash function."
    return $remoteFiles 
}

##############################################################################
# Running main
import-module (Join-Path $PSScriptRoot '..\lib\logs.psm1')
Main
# End of script
##############################################################################
