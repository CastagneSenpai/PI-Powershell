# Exemples d'utilisation :
# Write-Log -v_LogFile $LogPathFile -v_LogLevel SUCCESS -ConsoleOutput -v_Message "Ceci est un message de succès"
# Write-Log -v_LogFile $LogPathFile -v_LogLevel INFO -ConsoleOutput -v_Message "Ceci est un message d'information"
# Write-Log -v_LogFile $LogPathFile -v_LogLevel WARN -ConsoleOutput -v_Message "Ceci est un message d'avertissement"
# Write-Log -v_LogFile $LogPathFile -v_LogLevel ERROR -ConsoleOutput -v_Message "Ceci est un message d'erreur"
# Write-Log -v_LogFile $LogPathFile -v_LogLevel DEBUG -ConsoleOutput -v_Message "Ceci est un message de debug"
# Write-EmptyLine -v_LogFile $LogPathFile

Param(
    [string]$LogDir = (Join-Path (Get-Location) "Logs")
)

Function Write-Log {
    [CmdletBinding()]
    Param(
        [string]$v_Message,
        [string]$v_LogFile = (Join-Path -Path $LogDir -ChildPath ((Get-Date -Format yyyy-MM-dd) + "_Logs.txt")),
        [switch]$v_ConsoleOutput,
        [ValidateSet("SUCCESS", "INFO", "WARN", "ERROR", "DEBUG")]
        [string]$v_LogLevel = "INFO"
    )

    Begin {
        # Define log levels color
        $colorMap = @{
            "SUCCESS" = "Green"
            "INFO"    = "White"
            "WARN"    = "Yellow"
            "ERROR"   = "Red"
            "DEBUG"   = "Gray"
        }
        $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $color = $colorMap[$v_LogLevel]
        if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Force -Path $LogDir }
    }

    Process {
        if ($v_Message) {
            $logEntry = "[$timeStamp] [$v_LogLevel] :: $v_Message"

            try {
                if ($v_LogFile) {
                    Out-File -Append -FilePath $v_LogFile -InputObject $logEntry
                }

                if ($v_ConsoleOutput.IsPresent) {
                    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                    Write-Host $logEntry -ForegroundColor $color
                }
            }
            catch {
                Write-Error "Failed to write log: $_"
            }
        }
    }
}

Function Write-EmptyLine {
    [CmdletBinding()]
    Param(
        [string]$v_LogFile
    )

    Process {
        try {
            Write-Host ""  # Empty line

            if ($v_LogFile) {
                Out-File -Append -FilePath $v_LogFile -InputObject ""  # Empty line
            }
        }
        catch {
            Write-Error "Failed to write empty line: $_"
        }
    }
}

# Export functions
Export-ModuleMember -Function Write-Log
Export-ModuleMember -Function Write-EmptyLine
