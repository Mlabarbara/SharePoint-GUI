using namespace System.Windows.Forms
using namespace System.Drawing

function Get-SharePointBaseUrl { # this function will return the base URL of the SharePoint site from the configuration file
    $config = Get-EmailConfig
    if ($config -and $config.SharePointBaseUrl) {
        return $config.SharePointBaseUrl
    }
    else {
        Write-Error "SharePoint base URL not found in configuration"
        return $null
    }
}
