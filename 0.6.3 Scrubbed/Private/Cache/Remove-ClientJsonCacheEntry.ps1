using namespace System.Windows.Forms
using namespace System.Drawing

function Remove-ClientJsonCacheEntry { # Removes a client entry from the JSON cache using the -ClientName and -WhereIs parameters
    param (
        [Parameter(Mandatory=$true)]
        [string]$ClientName,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Folder", "TeamAmanda", "TeamJamie", "TeamKyle", "Redwood")]
        [string]$WhereIs
    )

    Write-Log "Removing client entry from JSON cache: $ClientName | WhereIs: $WhereIs"

    try {
        if($WhereIs -eq "Folder") {
            $jsonPath = Get-CacheFilePath -CacheType $WhereIs
        } else {
            $jsonPath = Get-CacheFilePath -CacheType "Subsite$WhereIs"
        }
        
        Write-Log "JSON path: $jsonPath"

        if (Test-Path $jsonPath) {
            $cacheData = Get-Content $jsonPath | ConvertFrom-Json

            if ($WhereIs -eq "Folder") {
                # Handle Folder cache
                $cacheData.General = @($cacheData.General | Where-Object { $_.ClientName -ne $ClientName })
                $cacheData.ResearchClients.TeamAmanda = @($cacheData.ResearchClients.TeamAmanda | Where-Object { $_.ClientName -ne $ClientName })
                $cacheData.ResearchClients.TeamJamie = @($cacheData.ResearchClients.TeamJamie | Where-Object { $_.ClientName -ne $ClientName })
            } else {
                # Handle TeamAmanda or TeamJamie or TeamKyle cache
                $cacheData = @($cacheData | Where-Object { $_.ClientName -ne $ClientName })
            }

            # Maintain alphabetical order
            if ($WhereIs -eq "Folder") {
                $cacheData.General = @($cacheData.General | Sort-Object -Property ClientName)
                $cacheData.ResearchClients.TeamAmanda = @($cacheData.ResearchClients.TeamAmanda | Sort-Object -Property ClientName)
                $cacheData.ResearchClients.TeamJamie = @($cacheData.ResearchClients.TeamJamie | Sort-Object -Property ClientName)
            } else {
                $cacheData = @($cacheData | Sort-Object -Property ClientName)
            }

            # Save the updated cache data
            $cacheData | ConvertTo-Json -Depth 10 | Set-Content $jsonPath

            Write-Log "Successfully removed client entry from JSON cache: $ClientName"
        } else {
            Write-Log "JSON file not found: $jsonPath" -Level "WARN"
        }
    } catch {
        Write-Log "Error removing client entry from JSON cache: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}
