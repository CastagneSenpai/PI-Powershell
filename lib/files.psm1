import-module (Join-Path $PSScriptRoot 'logs.psm1') 

Function Write-PITagData([System.Object] $PIPoint, [System.Object] $PIData, [string] $outputFolder, [string] $outputFile, [bool] $useUTC)
{
    #Writing of the data of the month in the file
     try
     {
         #Creation of the FileStream to allow writing to the current tag file
         $myDestinationForData = join-path $outputFolder $outputFile
         $fs = New-Object System.IO.FileStream $myDestinationForData ,'Append','Write','Read'
         $myStreamWriter =  New-Object System.IO.StreamWriter($fs)

         #Write to file
         ForEach ($line in $PIData)
         { 
            #Set point name
            $rPtName = $PIPoint.Point.Name
            
            #Set point value
            $rValue = $line.Value
            
            #Set value timestamp 
            if($useUTC){                
                $rtimestamp = $line.timestamp.ToUniversalTime().toString("o")
            }
            else{
                $rtimestamp = $line.timestamp.ToLocalTime().toString("o")
            }

            #Set quality status of value
            if($line.IsGood){
                $rIsGood = "Good Value"
            }
            else{
                $rIsGood = "No Good Value"
            }

            #Write line to file
            $myStreamWriter.WriteLine("$rPtName;$rtimestamp;$rValue;$rIsGood")   
        }

         #Closing the stream to the file     
         $myStreamWriter.Close() 
     }
     catch [System.Exception]
     {
         Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "Error in handling the output file $currentFile. [$($_.Exception.Message)]"
         continue
     }  
}

Function compress-ZipFolder([string] $outputFolder, [string] $outputFile)
{
    Try
    {
        $path = Join-Path $outputFolder $outputFile
        $destinationPath = $path + ".zip"
        Compress-Archive -Path $path -DestinationPath $destinationPath -Force -CompressionLevel Fastest
        Remove-Item -Path $path
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$currentFile : File successfully compressed."
    }
    catch [System.Exception]
    {
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "Error compressing file $outputFile. [$($_.Exception.Message)]"
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "The original text file has not been deleted."
        continue
    }
}

Function compress-ZipAllFolder([string] $outputFolder, [string] $name)
{
    Try
    {
		$path = $outputFolder + "\*.csv"
        $destinationPath = Join-Path $outputFolder $name
		
		
        Compress-Archive -Path $path -DestinationPath $destinationPath -Force -CompressionLevel Fastest
		
		Remove-Item -Path $path -Recurse
		
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel SUCCESS -v_ConsoleOutput -v_Message "$currentFile : File successfully compressed."
    }
    catch [System.Exception]
    {
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel ERROR -v_ConsoleOutput -v_Message "Error compressing file $outputFile. [$($_.Exception.Message)]"
        Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_ConsoleOutput -v_Message "The original text file has not been deleted."
        continue
    }
}

Export-ModuleMember -Function Write-PITagData
Export-ModuleMember -Function compress-ZipFolder
Export-ModuleMember -Function compress-ZipAllFolder