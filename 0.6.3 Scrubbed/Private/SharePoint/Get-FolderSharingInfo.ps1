using namespace System.Windows.Forms
using namespace System.Drawing

function Get-FolderSharingInfo { # TO USE: Get-FolderSharingInfo -siteUrl $siteUrl -folderPath $folderPath
    param (
        [string]$siteUrl,
        [string]$folderPath
    )
    try {
        #Write-Log "Connecting to site: $siteUrl"
Connect-PnPOnline
        Write-Log "Getting sharing info for folder: $folderPath"
        $sharingInfo = Get-PnPFolderSharingLink -Folder $folderPath
        
        if ($null -eq $sharingInfo) {
            throw "No sharing information found for folder: $folderPath"
        }
        Write-Log "Sharing link: $($sharingInfo.Link.WebUrl)"
        Write-Log "Granted identities: $($sharingInfo.GrantedToIdentitiesV2.User.Email -join ', ')"
        return @{
            SharingLink = $sharingInfo.Link.WebUrl
            GrantedIdentities = $sharingInfo.GrantedToIdentitiesV2.User.Email
        }
    }
    catch {
        Write-Log "Error in Get-FolderSharingInfo: $_" -Level "ERROR"
        return $null
    }
}
