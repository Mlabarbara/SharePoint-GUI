using namespace System.Windows.Forms
using namespace System.Drawing

function Get-CompanyUrl { # TO USE: Get-CompanyUrl -UrlType "Logo"
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Logo", "Support", "PrivacyPolicy")]
        [string]$UrlType
    )
    
    $config = Get-EmailConfig
    if ($config -and $config.CompanyUrls.$UrlType) {
        return $config.CompanyUrls.$UrlType
    }
    else {
        Write-Error "Company URL for '$UrlType' not found in configuration"
        return $null
    }
}
