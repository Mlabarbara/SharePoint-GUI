using namespace System.Windows.Forms
using namespace System.Drawing

function Format-EmailTemplate { # TO USE: Format-EmailTemplate -Template $template -Placeholders $placeholders
    param (
        [Parameter(Mandatory=$true)]
        [string]$Template,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Placeholders
    )
    
    foreach ($key in $Placeholders.Keys) {
        $placeholder = "{{$key}}"
        $value = $Placeholders[$key]
        $Template = $Template.Replace($placeholder, $value)
    }
    
    return $Template
}
