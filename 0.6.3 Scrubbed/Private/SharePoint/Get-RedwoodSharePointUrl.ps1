using namespace System.Windows.Forms
using namespace System.Drawing

function Get-RedwoodSharePointUrl { # this function will return the base URL of the SharePoint site from the configuration file
    $config = Get-EmailConfig
    if ($config -and $config.RedoodSharePointUrl) {
        return $config.RedoodSharePointUrl
    }
    else {
        Write-Error "Redood SharePoint URL not found in configuration"
        return $null
    }
}
