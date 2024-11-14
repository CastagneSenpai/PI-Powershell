import-module (Join-Path $PSScriptRoot 'logs.psm1')

Function Connect-PIServer([string]$PIServerHost, [int] $nbTry = 3)
{
    try 
    {
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$PIServerHost : Connection to the PI server in progress ..."
        $PIConnection = Connect-PIDataArchive -PIDataArchiveMachineName $PIServerHost -AuthenticationMethod Windows -ErrorAction Stop
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$PIServerHost : Connected successfully to the server."
        
        return $PIConnection
    }
    catch [System.Exception] 
    {
        if($nbTry -gt 0)
        {
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel WARN -v_ConsoleOutput -v_Message "$PIServerHost : Connection to server failed. [$($_.Exception.GetType().Name)]"
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$PIServerHost : Next try in 10 seconds (remaining $nbTry)."
            Start-Sleep -s 10
            Connect-PIServer -PIServerHost $PIServerHost -nbTry ($nbTry-1)
        }
        else
        {
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$PIServerHost : Connection to server failed. [$($_.Exception.Message)]"
            pause
            exit
        }        
    }
}

function Connect-AFDatabase {
    param (
        [string]$afServerName = "vmcegdidev001",
        [string]$afDBName = "Romain_Dev"
    )

    try {
        $afSystems = New-Object OSIsoft.AF.PISystems
        $afServer = $afSystems[$afServerName]
        $afServer.Connect()
        Write-Log -v_Message "Successfully connected to AF Server: $afServerName" -v_ConsoleOutput -v_LogLevel INFO
        $afDB = $afServer.Databases[$afDBName]

        if ($null -eq $afDB) {
            Write-Log -v_Message "Database $afDBName not found on AF Server $afServerName." -v_LogLevel ERROR -v_ConsoleOutput
        } else {
            Write-Log -v_Message "Successfully connected to AF Database: $afDBName" -v_ConsoleOutput -v_LogLevel INFO
        }
        return $afDB
    }
    catch {
        Write-Log -v_Message "Failed to connect to AF Server or Database: $_" -v_ConsoleOutput -v_LogLevel ERROR
        exit
    }
}

Function Test-PIConnection([System.Object] $PIConnection)
{
    while ($PIConnection.Connected -eq $false){
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_Message "$($PIConnection.CurrentRole.Name) : Connexion KO." 
        Start-Sleep -Milliseconds 5000  
    }
    #Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_Message "$($PIConnection.CurrentRole.Name) : Connexion OK." 
}

Function Get-PIpointSafe([String] $TagName, [System.Object] $PIConnection, [int] $nbTry = 2)
{            
    Test-PIConnection -PIConnection $PIConnection
    try 
    {
        $PIPoint = Get-PIpoint -Name $TagName -Connection $PIConnection -ErrorAction Stop
        if(!$PIPoint) { Write-Error -Message "Get-PIpoint for $point return null." -ErrorAction Stop }
        return $PIPoint
    }
    catch
    {
        if($nbTry -gt 0)
        {
            Get-PIpointSafe -TagName $TagName -PIConnection $PIConnection -nbTry ($nbTry-1)
        }
        else
        {
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$point : Issue using Get-PIpoint function - Tag extraction canceled. [$($_.Exception.Message)]"
            continue
        }
    }
}

Function Get-PIValuesSafe([System.Object] $PIPoint, [DateTime] $st, [DateTime] $et, [int] $nbTry = 3)
{
    Test-PIConnection -PIConnection $PIPoint.Point.Channel
    try
    {
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_Message "$($PIPoint.Point.Name) : Extraction $st >> $et in progress ..."
        $results = Get-PIValue -PIpoint $PIPoint -startTime $st -endTime $et -ErrorAction Stop | Select-Object Timestamp, Value, IsGood
        return $results
        
    }
    catch [System.Exception]
    {
        if($nbTry -gt 0)
        {
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_Message "$($PIPoint.Point.Name) : Retry Extraction $st >> $et."
            Get-PIValuesSafe -PIPoint $PIPoint -st $st -et $et -nbTry ($nbTry-1)
        }
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_Message "$($PIPoint.Point.Name) : Error with Extraction $st >> $et. Code will treat next month. [$($_.Exception.Message)]"
    }
}

Export-ModuleMember -Function Connect-PIServer
Export-ModuleMember -Function Connect-AFDatabase
Export-ModuleMember -Function Test-PIConnection
Export-ModuleMember -Function Get-PIValuesSafe
Export-ModuleMember -Function Get-PIpointSafe