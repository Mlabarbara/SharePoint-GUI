using namespace System.Windows.Forms
using namespace System.Drawing

function New-FolderClientObject { # Creates a new client object for the folder cache to ensure consistency
    param (
        [string]$clientName,
        [string[]]$clientEmails,
        [string]$url,
        [string]$sharingLink,
        [string]$type = "Folder",
        [string]$location
    )

    return [PSCustomObject]@{
        ClientName = $clientName
        ClientEmails = @($clientEmails)
        Url = $url
        SharingLink = $sharingLink -replace '\?email=.*$', ''
        Type = $type
        Location = $location
    }
}
