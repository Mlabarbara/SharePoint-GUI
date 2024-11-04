using namespace System.Windows.Forms
using namespace System.Drawing

function Test-ValidEmailFormat  {
    param (
        [string]$Email
    )
    
    # More comprehensive email regex that handles more edge cases
    $emailRegex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    
    # Additional checks
    if ([string]::IsNullOrWhiteSpace($Email)) {
        return $false
    }
    
    if ($Email.Length -gt 254) { # Maximum allowed email length
        return $false
    }
    
    return $Email -match $emailRegex
}
