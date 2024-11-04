using namespace System.Windows.Forms
using namespace System.Drawing

function Move-ClientFolderJson {
    param (
        [string]$clientName,
        [string]$newTeam,
        [string]$newSharingLink,
        [string]$newUrl,
        [string[]]$clientEmails
    )

    $jsonPath = Get-CacheFilePath -CacheType "Folder"
    Write-Log "Updating folder cache for client: $clientName | Moving to $newTeam"

    try {
        if (Test-Path $jsonPath) {
            $cacheData = Get-Content $jsonPath | ConvertFrom-Json
        } else {
            Write-Log "Cache file not found: $jsonPath" -Level "WARN"
            $cacheData = @{
                General = @()
                ResearchClients = @{}
            }
        }

        $clientInfo = $cacheData.General | Where-Object { $_.ClientName -eq $clientName }

        if ($clientInfo) {
            Write-Log "Client found in General section: $clientName"
            $cacheData.General = @($cacheData.General | Where-Object { $_.ClientName -ne $clientName })
        } else {
            Write-Log "Client not found in General section: $clientName" -Level "WARN"
        }

        $updatedClient = New-FolderClientObject -clientName $clientName -clientEmails $clientEmails `
                                                -url $newUrl -sharingLink $newSharingLink `
                                                -type "Folder" -location "Research Clients"

        if (-not $cacheData.PSObject.Properties['ResearchClients']) {
            $cacheData | Add-Member -NotePropertyName 'ResearchClients' -NotePropertyValue @{}
        }

        if (-not $cacheData.ResearchClients.PSObject.Properties[$newTeam]) {
            $cacheData.ResearchClients | Add-Member -NotePropertyName $newTeam -NotePropertyValue @()
        }
        $cacheData.ResearchClients.$newTeam += $updatedClient

        $cacheData.General = @($cacheData.General | Sort-Object -Property ClientName)
        $cacheData.ResearchClients.$newTeam = @($cacheData.ResearchClients.$newTeam | Sort-Object -Property ClientName)

        $cacheData | ConvertTo-Json -Depth 10 | Set-Content $jsonPath

        Write-Log "Updated folder cache for client: $clientName"
        Write-Log "Moved to Research Clients - $newTeam"
    } catch {
        Write-Log "Error updating folder cache: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}
