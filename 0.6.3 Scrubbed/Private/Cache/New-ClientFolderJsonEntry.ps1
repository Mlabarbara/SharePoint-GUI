using namespace System.Windows.Forms
using namespace System.Drawing

function New-ClientFolderJsonEntry {
    param (
        [string]$clientName,
        [string[]]$clientEmails,
        [string]$sharingLink,
        [string]$url,
        [string]$type = "Folder",
        [bool]$isResearchClient,  
        [string]$teamName
    )

    $jsonPath = Get-CacheFilePath -CacheType "Folder"
    Write-Log "Adding new client to folder cache: $clientName | $clientEmails | $sharingLink | $url | $type | $isResearchClient | $teamName"

    try {
        if (Test-Path $jsonPath) {
            $cacheData = Get-Content $jsonPath | ConvertFrom-Json
        } else {
            $cacheData = @{
                General = @()
                ResearchClients = @{
                    TeamAmanda = @()
                    TeamJamie = @()
                    TeamKyle = @()
                }
            }
        }
        # Initialize empty arrays for null team arrays
        if ($isResearchClient -and $null -eq $cacheData.ResearchClients.$teamName) {
            $cacheData.ResearchClients.$teamName = @()
        }
        $newClient = New-FolderClientObject -clientName $clientName -clientEmails $clientEmails `
                                            -url "/sites/secureclientupload/$url" -sharingLink $sharingLink `
                                            -type $type -location $(if ($isResearchClient) { "Research Clients" } else { "General" })

        if ($isResearchClient) {
            $targetArray = [System.Collections.ArrayList]::new($cacheData.ResearchClients.$teamName)
        } else {
            $targetArray = [System.Collections.ArrayList]::new($cacheData.General)
        }

        $insertIndex = 0
        while ($insertIndex -lt $targetArray.Count -and $targetArray[$insertIndex].ClientName -lt $clientName) {
            $insertIndex++
        }
        $targetArray.Insert($insertIndex, $newClient)

        if ($isResearchClient) {
            $cacheData.ResearchClients.$teamName = $targetArray
        } else {
            $cacheData.General = $targetArray
        }

        $cacheData | ConvertTo-Json -Depth 10 | Set-Content $jsonPath

        Write-Log "Successfully added new client to folder cache: $clientName"
    } catch {
        Write-Log "Error adding new client to folder cache: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}
