using namespace System.Windows.Forms
using namespace System.Drawing

# Helper function for subsite cache entries
function New-SubsiteClientObject {
    param (
        [string]$clientName,
        [string[]]$clientEmails,
        [string]$sharingLink,
        [string]$hasClientDocumentsFolder = "N",
        [double]$ageOfSubsiteMonths = 0
    )

    return [PSCustomObject]@{
        ClientName = $clientName
        ClientEmails = @($clientEmails)
        HasClientDocumentsFolder = $hasClientDocumentsFolder
        SharingLink = $sharingLink -replace '\?email=.*$', ''
        AgeOfSubsiteMonths = $ageOfSubsiteMonths
    }
}
