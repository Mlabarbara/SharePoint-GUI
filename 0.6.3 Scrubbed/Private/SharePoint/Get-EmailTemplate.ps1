using namespace System.Windows.Forms
using namespace System.Drawing

function Get-EmailTemplate { # TO USE: Get-EmailTemplate -TemplateName "ClientWelcome"
    param (
        [Parameter(Mandatory=$true)]
        [string]$TemplateName
    )

    Write-Log "Using template path: $script:templatePath"
    Write-Log "Using config path: $script:configPath"    
    Write-Log "Retrieving email template: $TemplateName"
    
    $config = Get-EmailConfig
    if ($config -and $config.Templates.$TemplateName) {
        $templateFile = Join-Path $script:templatePath ($config.Templates.$TemplateName + ".html")
        Write-Log "Template file path: $templateFile"
        if (Test-Path $templateFile) {
            $content = Get-Content $templateFile -Raw
            Write-Log "Template content retrieved. Length: $($content.Length)"
            return $content
        }
        else {
            Write-Log "Template file not found: $templateFile" -Level "ERROR"
            throw "Template file not found: $templateFile"
        }
    }
    else {
        Write-Log "Template '$TemplateName' not found in configuration" -Level "ERROR"
        throw "Template '$TemplateName' not found in configuration"
    }
}
