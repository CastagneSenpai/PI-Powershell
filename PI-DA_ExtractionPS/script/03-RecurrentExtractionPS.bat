cd /d %~dp0..
powershell.exe -ExecutionPolicy bypass -file ".\RecurrentExtractionPS.ps1" -PIServerHost "[PISERVERHOST]" -output "[OUTPUTPATH]" -doCompress -doCompressAll -noEmptyFile
cd script

