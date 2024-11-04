using namespace System.Windows.Forms
using namespace System.Drawing

function Get-ClientJsonCache {
    param (
        [string]$selectedTeam,
        [string]$selectedClient
    )
    
    Write-Log "Getting client info from JSON cache for team: $selectedTeam, client: $selectedClient"
    
    $jsonPath = Get-CacheFilePath -CacheType $(if ($selectedTeam -eq "TeamSales") { "Folder" } else { "Subsite$selectedTeam" })
    Write-Log "JSON path: $jsonPath"
    
    if (Test-Path $jsonPath) {
        try {
            Write-Log "Reading JSON file: $jsonPath"
            $clientData = Get-Content $jsonPath -Raw | ConvertFrom-Json
            Write-Log "Successfully parsed JSON data"
            
            if ($selectedTeam -eq "TeamSales") {
                if ($selectedClient) {
                    $clientInfo = @($clientData.General) + @($clientData.ResearchClients.TeamAmanda) + @($clientData.ResearchClients.TeamJamie) |
                                  Where-Object { $_.ClientName -eq $selectedClient }
                    Write-Log "Returning specific client data for TeamSales"
                    return $clientInfo
                } else {
                    Write-Log "Returning all TeamSales data"
                    return @{
                        General = @($clientData.General)
                        ResearchClients = @{
                            TeamAmanda = @($clientData.ResearchClients.TeamAmanda)
                            TeamJamie = @($clientData.ResearchClients.TeamJamie)
                            TeamKyle = @($clientData.ResearchClients.TeamKyle)
                        }
                    }
                }
            } else {
                if ($selectedClient) {
                    $clientInfo = @($clientData) | Where-Object { $_.ClientName -eq $selectedClient }
                    Write-Log "Returning specific client data for $selectedTeam"
                    return $clientInfo
                } else {
                    Write-Log "Returning all team-specific data for $selectedTeam"
                    return @($clientData)
                }
            }
        }
        catch {
            Write-Log "Error reading or parsing JSON: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
    }
    else {
        Write-Log "JSON file not found: $jsonPath" -Level "ERROR"
        throw "JSON file not found: $jsonPath"
    }
}
