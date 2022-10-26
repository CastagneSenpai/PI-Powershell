<#
.SYNOPSIS
This script find 5 verbose tags for each interface of PISC FUSION database and put them in the data collection in PI AF

.DESCRIPTION
This script find 5 verbose tags for each interface of PISC FUSION database and put them in the data collection in PI AF

.NOTES
Title: FindVerboseTags
Author: Romain CASTAGNE <romain.castagne@external.total.com>
Date : 21/10/2022
Version: 0.1
#> 

# Clean de l'interface 
clear-host

# Import Libraries
import-module (Join-Path $PSScriptRoot '..\lib\logs.psm1') 
import-module (Join-Path $PSScriptRoot '..\lib\Connection.psm1') 
[Reflection.Assembly]::LoadWithPartialName("OSIsoft.AFSDK") | Out-Null

# Log Repository
$v_LogPathfile = Join-Path $PSScriptRoot "logs\logs_$((get-date).toString('yyyy-MM-ddThh-mm-sstt')).txt"
$v_Reportfile = Join-Path $PSScriptRoot "logs\Reporting_$((get-date).toString('yyyy-MM-ddThh-mm-sstt')).txt"
Write-Log -v_LogLevel INFO -v_Logfile $v_LogPathfile -v_ConsoleOutput -v_Message "Script FindVerboseTags.PS1 start."

# Connection PISC
Write-Log -v_LogLevel INFO -v_Logfile $v_LogPathfile -v_ConsoleOutput -v_Message "Trying to access PISC Fusion..."
$AFServerName = "OPEPPA-WQPIHQ23"
$PISCDatabaseName = "PISC Fusion"
$PISCFusionServer = Get-AFServer -Name $AFServerName
Connect-AFServer -AFServer $PISCFusionServer | Out-Null
$PISCFusionDatabase = Get-AFDatabase -Name $PISCDatabaseName -AFServer $PISCFusionServer
Write-Log -v_LogLevel SUCCESS -v_Logfile $v_LogPathfile -v_ConsoleOutput -v_Message "PISC Fusion found."

#TO DELETE
$NumberOfTagsToCreate=0

# Get all interfaces elements to proceed
$RootElementName = "PISC"
$RootElement = Get-AFElement -name $RootElementName -AFDatabase $PISCFusionDatabase
$AffiliateElements = $RootElement | Select-Object Elements
foreach($CurrentAffiliateElement in $AffiliateElements.Elements)
{
    #Select child Elements starting with TEP* (Example : TEP Angola)
    if ($CurrentAffiliateElement.Name.StartsWith("TEP")){
        $AffiliateTrigramElement = $CurrentAffiliateElement | Select-Object Elements
        foreach($CurrentTrigramElement in $AffiliateTrigramElement.Elements)
        {
            #Select child Elements if the name is shorter or equal to 3 caractere (Example : LAD, CLV ...)  
            if($CurrentTrigramElement.Name.Length -le 3){
                $ServersElements = $CurrentTrigramElement | Select-Object Elements
                foreach($CurrentServerElement in $ServersElements.Elements)
                {
                    #Select child Element which represent Data Archive servers only (Example : AOEPTTA-APPIL01_DA)
                    if($CurrentServerElement.Name.EndsWith("_DA")){
                        $CurrentPIDANameLentgh = $CurrentServerElement.Name.Length

                        #Keep CurrentPIDAName : it will allow us to search the tags in each interfaces of it
                        if($CurrentServerElement.Name.EndsWith("_DMZ_DA")){
                            $CurrentPIDAName = $CurrentServerElement.Name.Substring(0, $CurrentPIDANameLentgh-7)
                        } else {
                            $CurrentPIDAName = $CurrentServerElement.Name.Substring(0, $CurrentPIDANameLentgh-3)
                        }

                        #Select Availability element
                        $AvailabilityElement = ($CurrentServerElement | Select-Object Elements).Elements[0]

                        #Select Interface Node elements
                        $InterfaceNodeElements = $AvailabilityElement.Elements
                        
                        #For each interface node element
                        forEach($CurrentInterfaceNodeElement in $InterfaceNodeElements)
                        {
                            $CurrentRedondancyElementName = ""
                            if($CurrentInterfaceNodeElement.Name.StartsWith("Redundancy"))
                            {
                                #If the element refers a Redundancy, get the child element instead 
                                $CurrentRedondancyElementName = "\" + $CurrentInterfaceNodeElement.Name
                                $CurrentInterfaceNodeElement = $CurrentInterfaceNodeElement.Elements[0]
                            }
                            #For each interface element
                            foreach($CurrentInterfaceElement in $CurrentInterfaceNodeElement.Elements)
                            {
                                Write-Log -v_LogLevel INFO -v_Logfile $v_LogPathfile -v_ConsoleOutput -v_Message "Processing interface $($CurrentInterfaceElement.Name) : `n`t`t`t`t`t`t`t`t(\\$AFServerName\$PISCDatabaseName\$RootElementName\$($CurrentAffiliateElement.Name)\$($CurrentTrigramElement.Name)\$($CurrentServerElement.Name)\$($AvailabilityElement.Name)$($CurrentRedondancyElementName)\$($CurrentInterfaceNodeElement.Name)\$($CurrentInterfaceElement.Name))"

                                #Get point source of the interface
                                $PointSourceAttribute = Get-AFAttribute -Name "Point Source" -AFElement $CurrentInterfaceElement
                                $PointSourceValue = $PointSourceAttribute.GetValue()

                                #Get location 1 of the interface
                                $InterfaceIDAttribute = Get-AFAttribute -Name "Interface ID" -AFElement $CurrentInterfaceElement
                                $InterfaceIDValue = $InterfaceIDAttribute.GetValue()

                                #Use ScanClassID = 1 to filter the list of tags to get (reduce performance issue)
                                $ScanClassID = 1

                                Write-log -v_LogLevel INFO -v_Logfile $v_LogPathfile -v_Message "Point source = $PointSourceValue, InterfaceID = $InterfaceIDValue"
                                
                                #Get Data Collection Monitoring Tags
                                $NumberOfTagsToFound = 0
                                
                                $DataCollectionMonitoringTags = Get-AFAttribute -Name "Data Collection Monitoring Tags" -AFElement $CurrentInterfaceElement
                                $DataCollectionMonitoringTagsAttributes = $DataCollectionMonitoringTags.Attributes
                                
                                #Count number empty attributes : to define the number of tags to find to fill the data collection tags in PI AF
                                foreach($TagAttribute in $DataCollectionMonitoringTagsAttributes){
                                    if($TagAttribute.GetValue().ToString().StartsWith("Tag Name is not specified"))
                                    {    
                                        #Tag Counter increase
                                        $NumberOfTagsToFound++
                                        $NumberOfTagsToCreate++
                                    }
                                }
                                if($NumberOfTagsToFound -gt 0)
                                {
                                    Write-Log -v_LogLevel WARN -v_Logfile $v_LogPathfile -v_Message "Number of tags to find :  $NumberOfTagsToFound"
                                    
                                    #Connect Data Archive
                                    $CurrentPIDACon = Connect-PIServer -PIServerHost $CurrentPIDAName
                                    
                                    #Get some tags for the interface
                                    $WhereClause = "Name:=* pointsource:=$PointSourceValue location1:=$InterfaceIDValue scan:=$ScanClassID"
                                    $PIPoints = Get-PIPoint -Attributes "tag" -Connection $CurrentPIDACon -WhereClause $WhereClause | Select-Object -First 500
                                    
                                    #Initialize a custom list, or reset it if new interface to proceed
                                    $TagsWithCountersTab = @()
                                    
                                    foreach($Tag in $PIPoints)
                                    {
                                        #Get 1 hour of values for the tags
                                        $PIValues = Get-PIValuesSafe -PIPoint $Tag -st (ConvertFrom-AFRelativeTime -RelativeTime "*30m") -et (ConvertFrom-AFRelativeTime -RelativeTime "*")
                                        
                                        # Create a custom Object to store the tag with the number of value associated
                                        $CurrentTagAndCounter = @{
                                            Counter = $PIValues.Count
                                            TagName = $Tag.Point.Name
                                        }    
                                        $TagsWithCountersTab += New-Object -TypeName PSCustomObject -Property $CurrentTagAndCounter                                        
                                    }
                                    #Select the 5 more verbose tags
                                    #TODO : Prendre 5 tags avec des noms de fonction différents
                                    #TODO : Skip les serveurs ou la connexion est KO (ex: NGEPLOS-APPIS02 DMZ) et mentionner de le faire manuellement. 
                                    $VerboseTags = $TagsWithCountersTab | Sort-Object Counter -Descending | Select-Object -First 5
                                    Write-Log -v_LogLevel WARN -v_Logfile $v_LogPathfile -v_ConsoleOutput -v_Message "(\\$AFServerName\$PISCDatabaseName\$RootElementName\$($CurrentAffiliateElement.Name)\$($CurrentTrigramElement.Name)\$($CurrentServerElement.Name)\$($AvailabilityElement.Name)$($CurrentRedondancyElementName)\$($CurrentInterfaceNodeElement.Name)\$($CurrentInterfaceElement.Name)) - `t Add this 5 tags : $($VerboseTags.TagName)"
                                    Write-Log -v_LogLevel INFO -v_Logfile $v_Reportfile -v_ConsoleOutput -v_Message "(\\$AFServerName\$PISCDatabaseName\$RootElementName\$($CurrentAffiliateElement.Name)\$($CurrentTrigramElement.Name)\$($CurrentServerElement.Name)\$($AvailabilityElement.Name)$($CurrentRedondancyElementName)\$($CurrentInterfaceNodeElement.Name)\$($CurrentInterfaceElement.Name)) - `t Add this 5 tags : $($VerboseTags.TagName)"
                                } else {
                                    Write-Log -v_LogLevel SUCCESS -v_ConsoleOutput -v_Logfile $v_LogPathfile -v_Message "5 tags are already mentionned for this element"
                                }
                            }   
                        }
                    }
                }
            }            
        }
    }    
}
Write-Log -v_Logfile $v_LogPathfile -v_Message "Number of tags to create :  $NumberOfTagsToCreate"

#End message
Write-Log -v_LogFile $v_LogPathfile -v_LogLevel INFO -v_Message "End of script."
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("End of script.", "FindVerboseTags.ps1", 0, 64)