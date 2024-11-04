using namespace System.Windows.Forms
using namespace System.Drawing

function Move-ClientFolder {# Moves client from General/TeamSales to Research Clients # TO USE: Move-ClientFolder -sourceSiteUrl $sourceSiteUrl -clientName $clientName -teamName $teamName
    param (
        [string]$sourceSiteUrl,
        [string]$clientName,
        [string]$teamName
    )
    
    try {
        Write-Log "Connecting to source site: $sourceSiteUrl"
Connect-PnPOnline
        
        $sourceFolder = Get-PnPFolder -Url "Shared Documents/General/TeamSales/$clientName"
        if ($null -eq $sourceFolder) {
            throw "Source folder not found: Shared Documents/General/TeamSales/$clientName"
        }
        
        $destinationFolderPath = "Shared Documents/Research Clients/$teamName"
        
        Write-Log "Moving folder from $sourceSiteUrl/$($sourceFolder) to $destinationFolderPath"
        Move-PnPFolder -Folder $sourceFolder -TargetFolder $destinationFolderPath
        
        Write-Log "Folder moved successfully"
        return $true
    }
    catch {
        Write-Log "Error in Move-ClientFolder: $_" -Level "ERROR"
        return $false
    }
}
