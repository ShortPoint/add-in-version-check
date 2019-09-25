# accept required parameters
<#
.SYNOPSIS
    .
.DESCRIPTION
    ShortPoint script to extract the list of site collection or subsites, url, site title, and ShortPoint Version installed to them .
.PARAMETER SharePointAdminUrl
    Specify the Url of your SharePoint Online Admin center. Url should be in format of https://<tenancyName>-admin.sharepoint.com .
.PARAMETER SharePointAdminUser
    Specify the SharePoint Online Administrator user name. Ex. admin@shortpoint.onmicrosoft.com .
.PARAMETER SharePointAdminPassword
    Specify the SharePoint Online Administrator password. 
.PARAMETER CSVExportFilePath
    The path with filename where you want Script to output the results. Ex. c:\ShortPoint\versions.csv.
.PARAMETER sharePointAssembliesPath
    We required following assemblies to make the script work;
        Microsoft.SharePoint.Client.Runtime.dll
        Microsoft.SharePoint.Client.dll
        Microsoft.Online.SharePoint.Client.Tenant.dll
        You can either download SharePoint Online client assemblies from MSDN OR you can use the one provided with the script zip in assemblies folder, here in this 
        zip file you will have to give path of the folder which contains above three assemblies.
.EXAMPLE
    C:\PS> ./VersionExtractorShortPoint.ps1 -SharePointAdminUrl "https://shortpoint-admin.sharepoint.com" -SharePointAdminUser "admin@shortpoint.onmicrosoft.com" -SharePointAdminPassword "PasswordOfAdminUser" -CSVExportFilePath "C:\ShortPoint\versions.csv" -sharePointAssembliesPath "C:\ShortPoint\VersionExtractor\Assemblies"
.NOTES
    Author: ShortPoint
    Date:   July 25, 2018    
#>
param(
[Parameter(Mandatory=$true,ValueFromPipeline=$true, HelpMessage="SharePoint Online admin site url. Ex: https://shortpoint-admin.sharepoint.com")]$SharePointAdminUrl, 
# [Parameter(Mandatory=$true,ValueFromPipeline=$true, HelpMessage="SharePoint Online admin user name.")][string]$SharePointAdminUser, 
# [Parameter(Mandatory=$true,ValueFromPipeline=$true, HelpMessage="SharePoint Online admin password.")][string]$SharePointAdminPassword, 
[Parameter(Mandatory=$true,ValueFromPipeline=$true, HelpMessage="Output csv path where you want to export the csv file having details of ShortPoint versions")][string]$CSVExportFilePath,
[Parameter(Mandatory=$true,ValueFromPipeline=$true, HelpMessage="Path of folder where the SharePoint Online assemblies are located.")][string]$sharePointAssembliesPath)

Add-Type -Path ($sharePointAssembliesPath + "\Microsoft.SharePoint.Client.Runtime.dll")
Add-Type -Path ($sharePointAssembliesPath + "\Microsoft.SharePoint.Client.dll")
Add-Type -Path ($sharePointAssembliesPath + "\Microsoft.Online.SharePoint.Client.Tenant.dll")

# Specify tenant admin and site URL
$global:outputExcelContent = @()

# Handling throttling
$Global:_retryCount = 1000
$Global:_retryInterval = 10

# Global settings
$host.Runspace.ThreadOptions = "ReuseThread"

# get list of all subsites
function Get-SPOSubWebs{ 
        Param( 
        [Microsoft.SharePoint.Client.ClientContext]$Context, 
        [Microsoft.SharePoint.Client.Web]$RootWeb 
        ) 

        # lets get the ShortPoint version for current web
        CheckAndGet-ShortPointVersion -Context $Context -Web $RootWeb -Site $null

        $Webs = $RootWeb.Webs 
        $Context.Load($Webs) 
        for($retryAttempts=0; $retryAttempts -lt $Global:_retryCount; $retryAttempts++){
             Try{
                    $Context.ExecuteQuery()
                    break
                }
             Catch [system.exception]{
             if($_.Exception.Response.StatusCode -eq 429)
              {
                Start-Sleep -s $Global:_retryInterval
              }
              else {
                break
              }
             }
         } 
 
        ForEach ($sWeb in $Webs) 
        { 
            Write-Output $sWeb.Title 
            Get-SPOSubWebs -RootWeb $sWeb -Context $Context 
        } 
} 

function CheckAndGet-ShortPointVersion {
Param( 
        [Microsoft.SharePoint.Client.ClientContext]$Context, 
        [Microsoft.SharePoint.Client.Web]$Web,
        [Microsoft.SharePoint.Client.Site]$Site
        ) 
        
        if($Site -ne $null)
        {
            $customUserActions = $Site.UserCustomActions
            $webId = $Site.Id
            $WebTitle = $Web.Title
            $WebUrl = $Site.Url
        }
        else
        {
            $customUserActions = $Web.UserCustomActions
            $webId = $Web.Id
            $WebTitle = $Web.Title
            $WebUrl = $Web.Url
        }
        #$Context.Load($customUserActions, 'Include(ScriptBlock, Sequence)') 
        $Context.Load($customUserActions)
        for($retryAttempts=0; $retryAttempts -lt $Global:_retryCount; $retryAttempts++){
             Try{
                    $Context.ExecuteQuery()
                    break
                }
             Catch [system.exception]{
                      if($_.Exception.Response.StatusCode -eq 429)
                      {
                        Start-Sleep -s $Global:_retryInterval
                      }
                      else {
                        break
                      }
             }
         }
 
        ForEach ($customUserAction in $customUserActions) 
        { 
            if(($customUserAction.Sequence -eq 200 -or $customUserAction.Sequence -eq 300) -and $customUserAction.Description -eq "ShortPoint.ScriptLink")
            {
                
                $scriptBlockText = $customUserAction.ScriptBlock
                $shortPointVersionVar = ([regex]::match( $scriptBlockText , "([\`"'])(\\?.)*?\1" )).value
                if($shortPointVersionVar -ne "")
                {
                    if($customUserAction.Sequence -eq 200)
                    {
                            $Scope = "Site collection"
                    }
                    else
                    {
                            $Scope = "Install per site"
                    }
                    Write-Host "ShortPoint Version for site " + $WebUrl + " is " + $shortPointVersionVar.Trim('"')
                    $SPTempObj = New-Object -TypeName PSObject -Property @{
                       ID = $webId 
                       Title = $WebTitle
                       Url = $WebUrl
                       Scope = $Scope
                       ShortPointVersion = $shortPointVersionVar.Trim('"')
                                              
                    }
                    # append temp object to global ShortPoint
                    $global:outputExcelContent += $SPTempObj
                }
            }
        } 

}
# Definition of the function that gets the list of site collections in the tenant using CSOM 
function Get-SPOTenantSiteCollections 
{ 
    param ($sSiteUrl,$sUserName,$sPassword) 
    try 
    {     
        Write-Host "----------------------------------------------------------------------------"  -foregroundcolor Green 
        Write-Host "Getting the Tenant Site Collections" -foregroundcolor Green 
        Write-Host "----------------------------------------------------------------------------"  -foregroundcolor Green 
      
      
        # SPO Client Object Model Context 
        Write-Host "Initializing spoCtx " + $sSiteUrl -foregroundcolor Yellow 
        $spoCtx = New-Object Microsoft.SharePoint.Client.ClientContext($sSiteUrl)
        $spoCtx.RequestTimeOut = 5000*10000  
        $spoCredentials = Get-Credential   
        Write-Host "Initializing spoCredentials" -foregroundcolor Yellow
        $spoCtx.Credentials = $spoCredentials 

        Write-Host "Connecting SPO service" -foregroundcolor Yellow
        #$SPOServiceCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sUsername, $sPassword 
        Connect-SPOService -Url $sSiteUrl -Credential $spoCredentials 
        

        $spoTenantSiteCollections = Get-SPOSite
        $startIndex = 0

             Write-Host "loading spoTenantSiteCollections" -foregroundcolor Yellow
                for($retryAttempts=0; $retryAttempts -lt $Global:_retryCount; $retryAttempts++){
                     Try{
                            $spoCtx.ExecuteQuery()
                            break
                        }
                     Catch [system.exception]{
                              if($_.Exception.Response.StatusCode -eq 429)
                              {
                                Start-Sleep -s $Global:_retryInterval
                              }
                              else {
                                break
                              }
                     }
                 }       
                    
            # We need to iterate through the $spoTenantSiteCollections object to get the information of each individual Site Collection 
            foreach($spoSiteCollection in $spoTenantSiteCollections){ 
                    Try{
                    Write-Host "Site Collection: " $spoSiteCollection.Url 
                    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($spoSiteCollection.Url)
                    $ctx.RequestTimeOut = 5000*10000  
                    $ctx.Credentials = $spoCredentials
                    for($retryAttempts=0; $retryAttempts -lt $Global:_retryCount; $retryAttempts++){
                        Try{
                              $ctx.ExecuteQuery()
                              break
                         }
                         Catch [system.exception]{
                              if($_.Exception.Response.StatusCode -eq 429)
                              {
                                Start-Sleep -s $Global:_retryInterval
                              }
                              else {
                                break
                                       }
                     }
                }
                    $Web = $ctx.Web 
                    $Site = $ctx.Site
                    $ctx.Load($Web)
                    $ctx.Load($Site)
                    for($retryAttempts=0; $retryAttempts -lt $Global:_retryCount; $retryAttempts++){
                        Try{
                              $ctx.ExecuteQuery()
                              break
                         }
                         Catch [system.exception]{
                            if($_.Exception.Response.StatusCode -eq 429)
                              {
                                Start-Sleep -s $Global:_retryInterval
                              }
                              else {
                                break
                              }
                        }
                     }  
            
                    # get the #ShortPoint version from Site collection
                    CheckAndGet-ShortPointVersion -Context $ctx -Site $Site -Web $Web

                    # get subsites recursivly 
                    Get-SPOSubWebs -Context $ctx -RootWeb $Web
                    }
                    Catch [system.exception]{
                     $SPTempObj = New-Object -TypeName PSObject -Property @{
                       ID = "" 
                       Title = ""
                       Url = $spoSiteCollection.Url
                       Scope = "Check Failed"
                       ShortPointVersion = "Check Failed"
                                              
                    }
                    # append temp object to global ShortPoint
                    $global:outputExcelContent += $SPTempObj
                    }
                }
 
        for($retryAttempts=0; $retryAttempts -lt $Global:_retryCount; $retryAttempts++){
         Try{
              $spoCtx.Dispose() 
              break
            }
            Catch [system.exception]{
                              if($_.Exception.Response.StatusCode -eq 429)
                              {
                                Start-Sleep -s $Global:_retryInterval
                              }
                              else {
                                break
                              }
                     }
                     }
        
    } 
    catch [System.Exception] 
    { 
        write-host -f red $_.Exception.ToString()    
    }     
} 

# Get list of site collections
Get-SPOTenantSiteCollections -sSiteUrl $SharePointAdminUrl 

# output the file to local path
$global:outputExcelContent |  Export-Csv -path $CSVExportFilePath -NoTypeInformation
