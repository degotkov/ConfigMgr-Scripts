 <#
    .SYNOPSIS
    Returns a list of maintenance windows for each collection located in the specified ConfigMgr (SCCM) folder.
    .DESCRIPTION
    The script uses the function Get-SCCMCollectionsInFolder shared by RaRaRatchet https://gist.github.com/RaRaRatchet/c8c4b113be4250db5dc18663fb24412e 
    and Powershell module ConfigurationManager.psd1 installed with ConfigMgr (SCCM) console.
  #>
Param([Parameter(Mandatory=$true)][string]$SiteCode,
    [Parameter(Mandatory=$true)][string]$FolderName,
    [Parameter(Mandatory=$true)][string]$SiteServer)

Function Get-SCCMCollectionsInFolder
{
  <#
    .SYNOPSIS
    Returns all the collections located in the specified folder ID.
    .DESCRIPTION
    Connects to the specified site server and retrieves the details of the specified folder as output that in term can be used for other functions.
    This function is usable for Device or User collections any other items will need different WMI queries and these would be best added to a seperate function.
   
    .EXAMPLE
    Get-SCCMCollectionsInFolder -FolderID <id of your folder>
 
    .EXAMPLE
    Get-SCCMCollectionsInFolder -FolderID <id of your folder> -SiteServer mysiteserver.example.com
 
    .PARAMETER FolderID
    This parameter is the folder ID (can be gathered using a different function Get-SCCMFolderDetail)
    .PARAMETER FolderType
    The FolderType parameter is used to specify the folder type (5000 for Device Collections or 5001 for User Collections)
    .PARAMETER SiteServer
    The SiteServer parameter contains the name of the site server that can provide the collections contained below.
    .PARAMETER SiteCode
    The SiteCode parameter is optional and if not provided automatically retrieved from the specified site server.
    .PARAMETER Full
    This is an optional and determines that all collection fields need to be gathered from the site server this will include member count etc.
    
    .AUTHOR
    RaRaRatchet
    
    .LINK
    https://gist.github.com/RaRaRatchet/c8c4b113be4250db5dc18663fb24412e
  #>
  Param
  (
    [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
    [string]$FolderID, 
    [Parameter(Mandatory=$False)]
    [string]$FolderType = "5000",
    [Parameter(Mandatory=$False)]
    [string]$SiteServer = "mysiteserver.example.com",
    [Parameter(Mandatory=$false)]
    [string]$SiteCode = (Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer).SiteCode,
    [Parameter(Mandatory=$False)]
    [switch]$Full = $false
  )
  Begin
  {
    Write-Verbose "SCCM Site Server                   : $($SiteServer)" 
    Write-Verbose "SCCM Site code                     : $($SiteCode)"
    Write-Verbose "SCCM Folder ID                     : $($FolderID)"
  }
  Process
  {
    Switch ($FolderType)
    {
      "5000" {$SCCMCollectionType = "2"}
      "5001" {$SCCMCollectionType = "1"}
      default {$SCCMCollectionType = "2"}
    }
    Write-Verbose "SCCM Collection Type               : $($SCCMCollectionType)"
    $FolderDetails = Get-WmiObject -Namespace "root\SMS\site_$($SiteCode)" -Query "Select * from SMS_ObjectContainerNode where`
    ContainerNodeID='$($FolderID)'" -ComputerName $SiteServer
    If ($Full)
    {
      Write-Verbose $FolderDetails
    }
    Else
    {
      Write-Verbose "SCCM Folder Name                   : $($FolderDetails.Name)"
    }
    $SCCMCollectionQuery ="select Name,CollectionID from SMS_Collection where CollectionID is in(select InstanceKey from SMS_ObjectContainerItem `
    where ObjectType='$($FolderType)' and ContainerNodeID='$FolderID') and CollectionType='$($SCCMCollectionType)'"
    If ($Full)
    {
      $SCCMCollectionQuery ="select * from SMS_Collection where CollectionID is in(select InstanceKey from SMS_ObjectContainerItem`
      where ObjectType='$($FolderType)' and ContainerNodeID='$FolderID') and CollectionType='$($SCCMCollectionType)'" 
    }
    $CollectionsInSpecficFolder = Get-WmiObject -Namespace "root\SMS\site_$($SiteCode)" -Query $SCCMCollectionQuery -ComputerName $SiteServer
    If ($VerbosePreference -eq "continue")
    {
      ForEach ($Collection in $CollectionsInSpecficFolder)
      {
        Write-Verbose "SCCM Collection Name               : $($Collection.name)"
        Write-verbose $Collection
      }
    }
  }
  End
  {
    return $CollectionsInSpecficFolder
  }
}

Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction SilentlyContinue
Set-Location "$($SiteCode):\" -ErrorAction SilentlyContinue

$Collections = Get-SCCMCollectionsInFolder -FolderID (Get-CMFolder -Name $FolderName).ContainerNodeID -SiteServer $SiteServer | sort name
foreach ($Collection in $Collections) {
    Write-Host $Collection.Name -ForegroundColor Cyan
    Get-CMMaintenanceWindow -CollectionId $Collection.CollectionID | ft name,description -AutoSize
    }
