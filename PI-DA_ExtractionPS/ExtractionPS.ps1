<#
.SYNOPSIS
Extract data from tags PI on a selected period.
 
 
.DESCRIPTION
Extract data from tags PI on a selected period.
 
 
.PARAMETER PIServerHost
[Mandatory] Name of the PI Server to extract tags.
 
 
.PARAMETER startTime
[Mandatory] Start time of extraction.


.PARAMETER endTime
[Mandatory] End time of extraction.


.PARAMETER timeZone
[Optional] UTC option (localtime as default).


.PARAMETER doCompress
[Optional] Compression option of the output files.


.PARAMETER doCompressAll
[Optional] Compression option of the output files in one zip.


.PARAMETER noEmptyFile
[Optional] No empty files generated option.
 
.NOTES
Title: ExtractionPS
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Version: 0.1
#> 


Param(
[Parameter(Mandatory=$True, HelpMessage="Name of the PI Server to extract tags (exemple : PI-CENTER-HQ)")][String]$PIServerHost,
[Parameter(Mandatory=$True, HelpMessage="Start time of extraction (Format : yyyy-MM-ddThh:mm:ss)")][String]$startTime,
[Parameter(Mandatory=$True, HelpMessage="End time of extraction (Format : yyyy-MM-ddThh:mm:ss)")][String]$endTime,
[Parameter(Mandatory=$false, HelpMessage="Output folder, default is output\ ")][String]$output,
[Parameter(Mandatory=$False, HelpMessage="Option to select <UTC> instead of <local> time")][Switch]$useUTC,
[Parameter(Mandatory=$False, HelpMessage="Option to compress the files after extraction ?")][Switch]$doCompress,
[Parameter(Mandatory=$False, HelpMessage="Option to compress the files after extraction in one zip?")][Switch]$doCompressAll,
[Parameter(Mandatory=$False, HelpMessage="Option to generate empty files ?")][Switch]$noEmptyFile
)
Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
import-module (Join-Path $PSScriptRoot '..\lib\logs.psm1') 
import-module (Join-Path $PSScriptRoot '..\lib\connection.psm1')
import-module (Join-Path $PSScriptRoot '..\lib\files.psm1')

if (!$output.Length){
    $outputFolder = Join-Path $PSScriptRoot "output\"
}else{$outputFolder= $output}

$sourceFile = Join-Path $PSScriptRoot "input\points.txt"
[datetime]$dateStartTime = $startTime
[datetime]$fixedDateStartTime = $startTime
[datetime]$dateEndTime = $endTime
$TotalTimeSpan = (New-TimeSpan -start $dateStartTime -end $dateEndTime)
$TotalTimeSpanInMinutes = $TotalTimeSpan.TotalMinutes
$v_LogPathfile = Join-Path $PSScriptRoot "logs\logs_$((get-date).toString('yyyy-MM-ddThh-mm-sstt')).txt"
$iTagCourant = 0 #Tag counter

#Loading tags in the input file
$points = Get-Content $sourceFile
$nbTagsTotal =  ($points | measure-object -line).Lines

#UTC Management
if($useUTC)
{
    $dateStartTime = $dateStartTime.ToUniversalTime()
    $dateEndTime = $dateEndTime.ToUniversalTime()
}

#Connection to the PI server
$PIConnection = Connect-PIServer -PIServerHost $PIServerHost

#TAG EXTRACTION
ForEach ($point in $points)
{
    #Increment of the tag counter
    $iTagCourant++  
    $continue = 1
    [datetime]$dateStartTime = $startTime

    Write-Progress -Activity ("PI-DA_ExtractionPS - Extraction of $nbTagsTotal tags in $PIServerHost server") -Status "Processing" -id 1 -PercentComplete ($iTagCourant/$nbTagsTotal*100)
    Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "$point ($iTagCourant/$nbTagsTotal) - Tag processing ..."
    
    $pt = Get-PIpointSafe -TagName $point -PIConnection $PIConnection
    Select-Object -Property timestamp, Value
    
    #Output file for the current tag
    $pointDisplay = $point.Replace("/", "_")
    $currentFile = "Extraction(" +$pointDisplay+ ")_" + $dateStartTime.toString('yyyy-MM-ddThh-mm-sstt') + "__" + $dateEndTime.toString('yyyy-MM-ddThh-mm-sstt') +".csv"
   
    #Retrieving data from the current tag
    while($continue)
    {
        $RemainingTimeSpan = (New-TimeSpan -start $dateStartTime -end $dateEndTime)
        $RemainingTimeSpanInMinutes = $RemainingTimeSpan.TotalMinutes
        Write-Progress -Activity ("Processing tag $point ($iTagCourant/$nbTagsTotal)") -Status "Extraction in progress ($fixedDateStartTime >> $dateEndTime) ..." -ParentId 1 -PercentComplete (($TotalTimeSpanInMinutes-$RemainingTimeSpanInMinutes)/$TotalTimeSpanInMinutes*100)

        $Date_Mid = $dateStartTime.AddDays(1)
        
        if($dateEndTime -gt $Date_Mid) #Process the current month
        {
            $results = Get-PIValuesSafe -PIPoint $pt -st $dateStartTime -et $Date_Mid
            $dateStartTime = $Date_Mid
        }
        else  #Process the last month
        {
            $results = Get-PIValuesSafe -PIPoint $pt -st $dateStartTime -et $dateEndTime
            $continue = 0  #No more data to get for the current tag
            Write-Progress -Activity ("Processing tag $point ($iTagCourant/$nbTagsTotal)") -Status "Extraction in progress ($fixedDateStartTime >> $dateEndTime) ..." -ParentId 1 -PercentComplete 100
        }

        #Create file if it's not empty and the option is set
        if(!$noEmptyFile -or $results){

            #Write current month data into the output file
            Write-PITagData -PIPoint $pt -PIData $results -outputFolder $outputFolder -outputFile $currentFile -useUTC $useUTC
        
            #Zip file and delete text file
            if(-not $continue){
                Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_Message "$point : Extraction complete."
                if($doCompress){
                    compress-ZipFolder -outputFolder $outputFolder -outputFile $currentFile
                }
            }
        }
    }
}

#Zip file and delete text file
if($doCompressAll){
    $name = "Compression(" + $dateStartTime.toString('yyyy-MM-ddThh-mm-sstt') + ").zip"
	compress-ZipAllFolder -outputFolder $outputFolder -name $name
}




Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "End of extraction."
[System.Windows.Forms.MessageBox]::Show("End of extraction.", "PI-DA_ExtractionPS", 0, 64)