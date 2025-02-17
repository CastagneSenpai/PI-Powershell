setlocal

@echo off
powershell.exe -executionpolicy bypass -noprofile -Command "& 'D:\PI\Applications\EventFrameBackfill\BackfillAnalysis.ps1'"

endlocal