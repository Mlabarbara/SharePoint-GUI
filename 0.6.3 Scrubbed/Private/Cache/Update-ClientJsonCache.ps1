using namespace System.Windows.Forms
using namespace System.Drawing

function Update-ClientJsonCache {
    param (
        [string]$teamName,
        [string]$clientName,
        [string[]]$email,
        [string]$sharingLink,
        [switch]$isNewSubsite
    )

    $jsonPath = Get-CacheFilePath -CacheType "Subsite$teamName"
    Write-Log "Updating JSON cache for client: $clientName | $email"

    try {
        if (Test-Path $jsonPath) {
            $clientData = Get-Content $jsonPath | ConvertFrom-Json
        } else {
            $clientData = @()
        }

        $clientToUpdate = $clientData | Where-Object { $_.ClientName -eq $clientName }

        if ($clientToUpdate) {
            $clientToUpdate.ClientEmails = @($email)
            $clientToUpdate.SharingLink = $sharingLink -replace '\?email=.*$', ''
            if ($isNewSubsite) {
                $clientToUpdate.HasClientDocumentsFolder = "Y"
                $clientToUpdate.AgeOfSubsiteMonths = 0
            }
        } else {
            $newClient = New-SubsiteClientObject -clientName $clientName -clientEmails $email `
                                                 -sharingLink $sharingLink `
                                                 -hasClientDocumentsFolder $(if ($isNewSubsite) { "Y" } else { "N" }) `
                                                 -ageOfSubsiteMonths 0
            $clientData += $newClient
        }

        $sortedClientData = @($clientData | Sort-Object -Property ClientName)

        $sortedClientData | ConvertTo-Json -Depth 10 | Set-Content $jsonPath

        Write-Log "Updated JSON cache for client: $clientName"
        Write-Log "Trimmed sharing link: $($sharingLink -replace '\?email=.*$', '')"
    } catch {
        Write-Log "Error updating JSON cache: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}
