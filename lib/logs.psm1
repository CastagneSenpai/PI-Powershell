# Exemples d'utilisation :
    # Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "Ceci est un message de succes"
    # Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "Ceci est un message d'information"
    # Write-Log -v_LogFile $v_LogPathfile -v_LogLevel WARN -v_ConsoleOutput -v_Message "Ceci est un message d'avertissement"
    # Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "Ceci est un message d'erreur"
    # Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "Ceci est un message de debug"
    # Write-EmptyLine -v_LogFile $v_LogPathfile

Param(
 [string]$v_LogDir
)

Function Write-Log(
[string[]]$v_Message, 
[string]$v_Logfile, 
[switch]$v_ConsoleOutput, 
[ValidateSet("SUCCESS", "INFO", "WARN", "ERROR", "DEBUG")]
[string]$v_LogLevel) {
 
 If (!$v_LogLevel) { $v_LogLevel = "INFO" }

 switch ($v_LogLevel) {
  SUCCESS { $v_Color = "Green" } 
  INFO { $v_Color = "White" } 
  WARN { $v_Color = "Yellow" } 
  ERROR { $v_Color = "Red" } 
  DEBUG { $v_Color = "Gray" } 
 }

 if ($v_Message -ne $null -and $v_Message.Length -gt 0) { 

  $v_TimeStamp = [System.DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss")

  if ($v_Logfile -ne $null -and $v_Logfile -ne [System.String]::Empty) {
   Out-File -Append -FilePath $v_Logfile -InputObject "[$v_TimeStamp] [$v_LogLevel] :: $v_Message"
  }

  if ($v_ConsoleOutput -eq $true){
   Write-Host "[$v_TimeStamp] [$v_LogLevel] :: $v_Message" -ForegroundColor $v_Color 
  } 
 }
}

If ($v_LogDir -eq "") {[string]$v_LogDir = Get-Location}
$v_LogPathfile = $v_LogDir + "\" + (Get-Date -Format yyyy-MM-dd) + "_Logs.txt"


Function Write-EmptyLine(
    [string]$v_Logfile) {
        Write-Host "" #EmptyLine
        if ($v_Logfile -ne $null -and $v_Logfile -ne [System.String]::Empty) {
            Out-File -Append -FilePath $v_Logfile -InputObject "" #EmptyLine
        }

    } 


Export-ModuleMember -Function Write-Log
Export-ModuleMember -Function Write-EmptyLine