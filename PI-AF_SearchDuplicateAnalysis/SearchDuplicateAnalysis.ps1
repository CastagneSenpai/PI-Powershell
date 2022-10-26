<#
.SYNOPSIS
Search all attribute on an AF server which write on input tag.
 
 
.DESCRIPTION
Search all attribute on an AF server which write on input tag.

 
.PARAMETER PIServerHost
[Mandatory] Name of the PI Server to extract tags.
 
 
.PARAMETER startTime
[Mandatory] Start time of extraction.

 
.NOTES
Title: PI-AF_SearchDuplicateAnalysis
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Version: 0.1
#> 


Param(
[Parameter(Mandatory=$True, HelpMessage="AF Server to search")]$afServ,
[Parameter(Mandatory=$True, HelpMessage="PI Tag to search")][String]$piTag
)


#Clear-Host
import-module (Join-Path $PSScriptRoot 'lib\logs.psm1') 
$v_LogPathfile = Join-Path $PSScriptRoot "logs\logs_$((get-date).toString('yyyy-MM-ddThh-mm-ss')).txt"
$outputFolder = Join-Path $PSScriptRoot "output\find_$piTag.txt"

$afServ = Get-AFServer -Name $afServ

$afDBs = Get-AFDatabase -AFServer $afServ
foreach ($afDB in $afDBs)
{
  Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "--------------Searching AF Database $($afDB.Name)"
  $afElems = [OSIsoft.AF.Asset.AFElement]::FindElements($afDB, $null, "*", [OSIsoft.AF.AFSearchField]::Name, $true, [OSIsoft.AF.AFSortField]::Name, [OSIsoft.AF.AFSortOrder]::Ascending, $([int]::MaxValue))
  Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "  Total $($afElems.Count) Elements found."

  foreach ($afElem in $afElems)
  {
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "--------Searching $($afElem.Name)"
    foreach ($afAttr in $afElem.Attributes)
    {
      if ($afAttr.DataReferencePlugIn.Name -like "*PI Point*")
      {
        if ($afAttr.DataReference.PIPoint.Name -like "*" + $piTag + "*")
        {
          Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "Found $($afAttr.Name)! Path : $($afAttr.GetPath())"
          Write-Log -v_LogFile $outputFolder -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$($afAttr.GetPath()) -- $($afAttr.Name)"
        }
      }
    }#inside each Attribute Object
  }#inside each Element Object
} 
