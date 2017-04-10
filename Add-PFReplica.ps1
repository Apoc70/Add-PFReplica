<#
.SYNOPSIS
Add legacy public folder replicas for multiple legacy public folder servers recursively 
   
Thomas Stensitzki
	
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
Version 1.2, 2017-04-10

Ideas, comments and suggestions to support@granikos.eu 
 
.LINK  
More information can be found at http://scripts.granikos.eu 
	
.DESCRIPTION
	
This script adds multiple public folder servers to all public folders below a TopPublicFolder.

The script has been developed as part of an on-premise legacy public folder migration from Exchange 2007 to Exchange Server 2010.

The script waits a given timespan in seconds to let public folder hierarchy replication and replica backfill requests kick in.

It is assumed that the script is being run in an Exchange 2007 or Exchange 2010 server.

.NOTES 
Requirements 
- Windows Server 2008 R2 SP1+
- Exchange Server 2007, Exchange Server 2010  
- Exchange Management Shell

Revision History 
-------------------------------------------------------------------------------- 
1.0     Initial community release 
1.1     Fixes to run properly with Exchange 2010
1.2     Minor PowerShell fixes, no additional functionality
	
.PARAMETER ServersToAdd
String array containing the legacy public folders to add

.PARAMETER SecondsToPause
Seconds to pause between each step, this should be a reasonalbe timespan (default = 240)

.PARAMETER PublicFolderServer
The server name of the legacy public folder server to contact for changes

.PARAMETER TopPublicFolder
Name of the legacy top public folder
   
.EXAMPLE
Add replicas for SERVER01,SERVER02 to all sub folders of \COMMUNICATIONS\PR
.\Add-PFReplica.ps1 -ServersToAdd SERVER01,SERVER02 -PublicFolderServer SERVER01 -TopPublicFolder "\COMMUNICATIONS\PR"

#>
param(    
    [parameter(Mandatory=$true,HelpMessage='String array containing the legacy public folders to add')]
    [string[]]$ServersToAdd,
    [int]$SecondsToPause = 240,
    [parameter(Mandatory=$true,HelpMessage='The server name of the legacy public folder server to contact for changes')]
    [string]$PublicFolderServer,
    [parameter(Mandatory=$true,HelpMessage='Name of the legacy top public folder')]
    [string]$TopPublicFolder
)

Write-Host ('Fetching public folders with TopPublicFolder: {0}' -f $TopPublicFolder)

# Fetch top public folder sub folders
$publicFolders = Get-PublicFolder $TopPublicFolder -Recurse -ResultSize Unlimited

# Some count stuff
$pfCount = ($publicFolders | Measure-Object)
$i = 1
$skippedFolders = 0
$maintainedFolders = 0
$max = [int]$pfCount.Count

if($PublicFolders -ne $null) {

    Write-Host ('Public folder count: {0}' -f $pfCount.Count)
    Write-Host ('Script will wait {0} seconds after each public folder' -f $SecondsToPause)

    # Now let's iterate through the public folders
    foreach($folder in $publicFolders) {
        $folderName = ('{0}\{1}' -f $folder.ParentPath, $folder.Name)
        
        $action = ('Working on [{0}/{1}]: {2}' -f $i, $max, $folderName)
        $status = "Fetching public folder data"
        
        Write-Progress -Activity $action -Status $status -PercentComplete(($i/$max)*100)
        
        $tmpReplicaServer = @()
        
        # Fetch Public Folder to work on
        $pFolder = Get-PublicFolder $folderName
       
        $presentCount = 0
        
        Foreach($server in $ServersToAdd) {
        
            $database = Get-PublicFolderDatabase -Server $server -ErrorAction Stop
            
            if(!$pFolder.Replicas.Contains($database.Identity)) {
                $status = ('Adding {0}' -f $server.ToUpper())
                Write-Progress -Activity $action -Status $status -PercentComplete(($i/$max)*100)
                
                # Fetch PF Database            
                $database = Get-PublicFolderDatabase -server $server 
                
                # Add PF database to replica list
                $pFolder.Replicas += $database.Identity

                # Set replicas
                Set-PublicFolder -Server $PublicFolderServer -Identity $pFolder.Identity -Replicas $pFolder.Replicas
                
            }
            else {
                # Write-Host "$($server) already present"
                $presentCount++
            }
        }
        
        if($presentCount -eq ($ServersToAdd.GetUpperBound(0) +1)) {
            # Write-Host "Skip waiting period"
            $skippedFolders++
        }
        else {
            $maintainedFolders++
            $status = ('Wait {0} seconds' -f $SecondsToPause)
            Write-Progress -Activity $action -Status $status -PercentComplete(($i/$max)*100)
            Start-Sleep -Seconds $SecondsToPause
        }
        
        $i++
    }

    Write-Host ('Public folder skipped: {0}' -f $skippedFolders)
    Write-Host ('Public folders received a new replica configuration: {0}' -f $maintainedFolders)
    Write-Host 'Script finished!'
} 
else {
    Write-Host ('TopPublicFolder {0} does NOT exist.' -f $TopPublicFolder)
}