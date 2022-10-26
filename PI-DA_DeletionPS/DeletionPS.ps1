<#
.SYNOPSIS
This script delete tags (data and structure) listed in the input file.
 
 
.DESCRIPTION
This script delete tags (data and structure) listed in the input file.
 
 
.PARAMETER piServerHost
[MANDATORY] PI DA Server where the tags will be deleted.
  
 
.NOTES
Title: DeletionPS
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Version: 0.1
#> 


Param(
[Parameter(Mandatory=$True, HelpMessage="Name of the PI Server to delete tags (exemple : PI-CENTER-HQ)")][String]$piServerHost
)

import-module (Join-Path $PSScriptRoot '..\lib\Logs.psm1') 

#Initialization of variables and logs.
$timestamp=get-date -uFormat "%m%d%Y%H%M"
$v_LogPathfile = (Join-Path $PSScriptRoot "logs\logs_$timestamp.txt") 
$sourceFile = (Join-Path $PSScriptRoot "input\points.txt")
[DateTime] $startTime = (Get-Date -Year 2000 -Month 01 -Day 01 -Hour 00 -Minute 00)
[DateTime] $endTime = Get-Date
Clear-Host

#Loading tags in the input file + indicators
$points = Get-Content $sourceFile 
$currentTag = 0
$totalTags =  ($points | measure-object -line).Lines

#Connection to the PI server
try 
{
    $piConnection = Connect-PIDataArchive -PIDataArchiveMachineName $piServerHost -AuthenticationMethod Windows -ErrorAction Stop
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$piServerHost : Connected to the server."
}
catch [System.Exception] 
{
    #If KO connection, end of program
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "$piServerHost : Connection to server failed. [$($_.Exception.Message)]."
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "End of application processing."
    Pause
    exit
}

#User Validation
Read-Host -Prompt "Press <Enter> to validate the deletion of the tags : $points"

ForEach($point in $points)
{
    $currentTag++
    $currentTagExist = 1
    [DateTime] $currentDateTime = (Get-Date -Year 2000 -Month 01 -Day 01 -Hour 00 -Minute 00)
    $continue = 1

    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "$point ($currentTag/$totalTags) : Data suppression in progress..."

    #DATA DELETION MONTH PER MONTH
    While($currentDateTime.AddMonths(1) -lt $endTime)
    { 
        #Delete the current month 
        try
        {
            $piEvents = Get-PIValue -PointName $point -StartTime $currentDateTime -EndTime $currentDateTime.AddMonths(1) -Connection $piConnection -ErrorAction Stop
            Remove-PIValue -Event $piEvents -Connection $piConnection -ErrorAction Stop
        }
 
        #Current tag does not exist : Get-PIValue return an exception
        catch [OSIsoft.PI.Net.PIObjectNotFoundException]
        {   
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel WARN -v_ConsoleOutput -v_Message "$point : Tag does not exist on $piServerHost, it can't be deleted."
            $currentTagExist = 0
            break
        }
 
        #Current tag exist, but no data in the current month : Remove-PIValue return an exception because piEvents is null
        catch [System.Management.Automation.ParameterBindingException]
        {               
            #Treat next month
            #Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "$point : No data to suppress between $currentDateTime and $($currentDateTime.AddMonths(1))"
            continue 
        }
 
        #Current tag exist, but the PI Server is not licensed to access for passed date
        catch [OSIsoft.PI.Net.PIException]
        {
            #Write-Log -v_LogFile $v_LogPathfile -v_LogLevel DEBUG -v_ConsoleOutput -v_Message "$piServerHost ($currentDateTime): [$($_.Exception.Message)]"
            continue
        }
 
        finally
        {            
            $currentDateTime = $currentDateTime.AddMonths(1)
        }
    }
 
    if($currentTagExist)
    {
        
        $piEvents = Get-PIValue -PointName $point -StartTime $currentDateTime -EndTime $endTime -Connection $piConnection
        try {Remove-PIValue -Event $piEvents -Connection $piConnection} catch{}
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$point : Data well suppressed."
        
        #TAG STRUCTURE DELETION
        try
        {
            Remove-PIpoint -Connection $piConnection -Name $point -ErrorAction Stop
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$point : Tag successfully deleted."
        }
        catch [System.Exception]
        {
            Write-Log -v_LogFile $v_LogPathfile -v_LogLevel WARN -v_ConsoleOutput -v_Message "$point : Tag cannot be deleted. [$($_.Exception.Message)]."
        }
    }
    else
    {
        continue
    }
}

Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "End of suppression process."
Disconnect-PIDataArchive -Connection $piConnection | Out-Null
Read-Host -Prompt "Press <Enter> to quit DeletionPS."