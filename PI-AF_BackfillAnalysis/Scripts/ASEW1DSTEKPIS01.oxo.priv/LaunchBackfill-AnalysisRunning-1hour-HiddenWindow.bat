setlocal
@echo off

start /b powershell.exe -executionpolicy bypass -noprofile -WindowStyle Hidden -Command "Set-Variable -Name afServerName -Value 'ASEW1DSTEKPIS01.oxo.priv'; Set-Variable -Name afDBName -Value 'Test_Prd_Posting'; Set-Variable -Name afSDKPath -Value 'D:\OSISOFT\PIPC_x86\AF\PublicAssemblies\4.0\OSIsoft.AFSDK.dll'; Set-Variable -Name RecalculationLogFilePath -Value 'C:\ProgramData\OSIsoft\PIAnalysisNotifications\Data\Recalculation\recalculation-log.csv'; Set-Variable -Name DeltaStartInMinutes -Value 80; Set-Variable -Name DeltaEndInMinutes -Value 0; Set-Variable -Name AutomaticMode -Value $true; Set-Variable -Name StartAndStopAnalysis -Value $false; Set-Variable -Name CategoriesName -Value @('Autobackfill_First', 'Autobackfill_Last'); & 'D:\PI\Applications\EventFrameBackfill\BackfillAnalysis.ps1' -afServerName $afServerName -afDBName $afDBName -afSDKPath $afSDKPath -DeltaStartInMinutes $DeltaStartInMinutes -DeltaEndInMinutes $DeltaEndInMinutes -RecalculationLogFilePath $RecalculationLogFilePath -CategoriesName $CategoriesName -AutomaticMode $AutomaticMode -StartAndStopAnalysis $StartAndStopAnalysis"

endlocal