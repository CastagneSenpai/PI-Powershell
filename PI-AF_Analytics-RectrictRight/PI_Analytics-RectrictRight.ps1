#Powershell script to list all point source of a PI server 
#and create acceptance test file with PI Values for each point source 
#Output files :
# "SPINUP - Listing PointSource DA - $PIServerName" : Listing of pointsource on the PI Server
# "SPINUP - Recette PointSource DA - $PIServerName" : Listing of PI Values for each point source"

param(
	[Parameter(Mandatory=$true)]
	[string] $PIServerName)

# connect to PI Data Archive
$con = Connect-PIDataArchive -PIDataArchiveMachineName $PIServerName

# Set the list of attributes to retrieve. o
$AttribList = ("PtSecurity", "DataSecurity", "PointSource")  
write-output $AttribList

# Get a reference to your PI Point.  
$Points = Get-PIPoint -Connection $con -Name "*" -Attributes $AttribList  

#write-output $($Points.Attributes.pointsource)

[int]$i=1

foreach ($PIPoint in $Points) 
{
	Write-Host "$(Get-Date) [DEBUG] Doing tag '$($PIPoint.Point.Name)' $($PIPoint.Attributes)"
	if ($i % 100 -eq 0) {
		Write-Output "$(Get-Date) [DEBUG] Doing tag '$($PIPoint.Point.Name)' $($i)/$($Points.Count)"
	}
	
    #Write-Output $PIPoint.Attributes.pointsource

    if(($PIPoint.Attributes.pointsource -like "*_D0*") -or ($PIPoint.Attributes.pointsource -like "*_N0*"))
    {
       #Write-Output $PIPoint.Attributes.pointsource

        # Transform security descriptors  
        #<#$AttribPtSecurity = @{"PtSecurity" = "piadmin: A(R,W) | piadmins: A(R,W) | PIWorld: A() | PI_Administrators: A(R,W) | PI_Users_Read_Write: A(R,W) | PI_Users_Read_Only: A(R) | PI_Analytics: A(R) | PI_Applications: A(R,W) | PI_AssetFramework: A(R) | PI_Buffer: A(R,W) | PI_Interfaces: A(R) | PI_Notifications: A(R) | PI_PerfMon: A() | PI_Vision: A(R)| PIWorld: A()"}  
        #$AttribDtSecurity = @{"DataSecurity" = "piadmin: A(R,W) | piadmins: A(R,W) | PIWorld: A() | PI_Administrators: A(R,W) | PI_Users_Read_Write: A(R,W) | PI_Users_Read_Only: A(R) | PI_Analytics: A(R) | PI_Applications: A(R,W) | PI_AssetFramework: A(R) | PI_Buffer: A(R,W) | PI_Interfaces: A(R) | PI_Notifications: A(R) | PI_PerfMon: A() | PI_Vision: A(R)| PIWorld: A()"}  
        #>
        #$securityAttributes = @{
        #    "PtSecurity" = $($PIPoint.Attributes.ptsecurity).replace("PI_Analytics: A(R,W)", "PI_Analytics: A(R)")
        #    "DataSecurity" = $($PIPoint.Attributes.datasecurity).replace("PI_Analytics: A(R,W)", "PI_Analytics: A(R)")
        #}

        Write-Output "Modification of $($PIPoint.Point.Name) in progress..."

        # Send the change to the PI Server.  
        #<#Set-PIPoint -PIPoint $PIPoint -Attributes $AttribPtSecurity  
        #Set-PIPoint -PIPoint $PIPoint -Attributes $AttribDtSecurity #>

         Set-PIPoint -PIPoint $PIPoint -Attributes @{
            "ptsecurity" = $($PIPoint.Attributes.ptsecurity).Replace("PI_Analytics: A(r,w)", "PI_Analytics: A(r)")
            "datasecurity" = $($PIPoint.Attributes.datasecurity).Replace("PI_Analytics: A(r,w)", "PI_Analytics: A(r)")
        }

        Write-Output "Modification of $($PIPoint.Point.Name) done."

    }
	
	$i++
      
}