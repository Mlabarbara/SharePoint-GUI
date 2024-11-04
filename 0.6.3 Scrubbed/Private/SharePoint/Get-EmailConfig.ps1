using namespace System.Windows.Forms
using namespace System.Drawing

function Get-EmailConfig { # parses the JSON for email, used in Get-EmailConfig 
    if (Test-Path $script:configPath) {
        $config = Get-Content $script:configPath | ConvertFrom-Json
        return $config
    }
    else {
        Write-Error "Configuration file not found at $script:configPath"
        return $null
    }
}
