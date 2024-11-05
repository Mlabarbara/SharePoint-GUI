Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ModuleLogPath = "\\str-0111\Logs"
$script:ModuleVersion = "0.6.3-Local"
$script:LOGFILE = Join-Path $script:ModuleLogPath "$($env:USERNAME)--$script:ModuleVersion--$((Get-Date).ToString('M-dd-yy')).log"
$script:PNPPOWERSHELL_UPDATECHECK = 'off'
$script:LOCAL_ROOT = "$env:LOCALAPPDATA\SequoiaTax\$script:ModuleVersion"
$script:NETWORK_ROOT = "\\str-0111\MainMenu\$script:ModuleVersion"

# Config paths
$script:configPath = "$script:NETWORK_ROOT\config.json"
$script:templatePath = "$script:NETWORK_ROOT\Templates"
$script:cachePath = "$script:NETWORK_ROOT\cache"

# Import all functions
$Public = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -Recurse -ErrorAction SilentlyContinue)
$Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -Recurse -ErrorAction SilentlyContinue)


foreach ($import in @($Public + $Private)) {
    try {
        . $import.FullName
    }
    catch {
        Write-Error "Failed to import function $($import.FullName): $_"
    }
}

Initialize-Win32Functions
# Export public functions
Export-ModuleMember -Function * -Variable *

