Clear-host

$serverNamesList = 
"AOEPTTA-APPIL01",
"AOEPCLV-APPIC02",
"AOEPDAL-APPID02",
"AOEPGIR-APPIG01",
"AOEPPAZ-APPIP01",
"AREPAGU-APPI01",
"AREPBUE-APPI01",
"AREPCAL-APPI01",
"AREPRCU-APPI01",
"AREPLLY-APPI01",
"BNEPLUM-AP05",
"BOEPSRZ-APPI01",
"BREPRIO-APPI01",
"BREPLAP-APPI01",
"CGEPBIN-APPI04",
"DKEPEBJ-APPIHS1",
"GAEPPOG-PIS01",
"GBEPABZ-APMIS01",
"GBEPABZ-APPLK01",
"GBEPGRY-APMIS01",
"ITEPCOR-AP03",
"MMEPRGN-APPI01",
"MMEPYDN-APPI01",
"NGEPLOS-APPIS01",
"NGEPAKP-APPIS01",
"NGEPEGN-APPIS01",
"NGEPPHC-APPI01",
"NGEPAMQ-APPI01",
"NGEPOBG-AP02",
"NLEPDHG-PI01",
"NOEPSVG-APPI02",
"QAEPDOH-WPPI01",
"OPEPPA-WRPIAO01",
"OPEPPA-WRPIAR01",
"OPEPPA-WRPIBR01",
"OPEPPA-WRPICG01",
"OPEPPA-WRPIDK01",
"OPEPPA-WRPIGA01",
"OPEPPA-WRPIGB01",
"OPEPPA-WRPIIT01",
"OPEPPA-WRPING01",
"OPEPPA-WRPING02",
"OPEPPA-WRPINL01",
"OPEPPA-WRPIQA01"

foreach($serverName in $serverNamesList)
{
    # CONNECTION
    try
    {
        Write-Host "Connecting to" $serverName "..."
        $con = Connect-PIDataArchive -PIDataArchiveMachineName $serverName
        Write-Host "Successfully connected to" $serverName `n`
    }
    catch
    {
        Write-Host "Error connecting to" $serverName `n`
        continue
    }

    # REPLACE PI_ADMINISTRATOR IN PIADMIN
    if ($con.Connected)
    {
        Try
        {
            $MappingList = Get-PIMapping -Connection $con
            foreach ($Mapping in $MappingList)
            {
                if($Mapping.Identity -eq "PI_Administrators")
                {
                    Write-Host "Server" $serverName "- Mapping" $Mapping.PrincipalName "will be deleted" `n`
                    Read-Host "OK ?"

                    Remove-PIMapping -Connection $con -Name $Mapping.Name

                    Write-Host "Server" $serverName "- Mapping" $Mapping.PrincipalName "will be recreated with identity piadmin" `n`
                    Read-Host "OK ?"

                    Add-PIMapping -Connection $con -Identity piadmin -Name $Mapping.PrincipalName -PrincipalName $Mapping.PrincipalName -Description $Mapping.Description
                }
            }
            
        }
        Catch
        {
            Write-Host "Error treating mappings update for PI server" $serverName
            Read-Host "Press ENTER to continue ..."
        }
         
    }
}