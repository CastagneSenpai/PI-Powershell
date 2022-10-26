import-module (Join-Path $PSScriptRoot 'logs.psm1') 

Function Get-NumericalLength([String] $value)
{
    $v_StartIndex = 7
    $v_lg=1
    $v_continue = 1
    #Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "v_StartIndex = $v_StartIndex - v_lg = $v_lg - value = $value"
    while($v_continue)
    {
        if($value.substring($v_StartIndex,$v_lg) -match "^\d+$")
        {
            #Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "$($line.tag)- $($line.timestamp) - $v_StartIndex - $v_lg"
            $v_lg++
            $v_StartIndex++
        }
        else
        {
            $v_lg--
            $v_continue = 0
        }
    }
    #Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "v_lg = $v_lg"
            
    return $v_lg
}