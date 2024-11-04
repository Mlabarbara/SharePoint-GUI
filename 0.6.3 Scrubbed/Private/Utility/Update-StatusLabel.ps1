using namespace System.Windows.Forms
using namespace System.Drawing

function Update-StatusLabel { # Displays a status message to the User. TO USE: Update-StatusLabel -Label $label -Form $form -Status "Message to display"
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Label]$Label,
        
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$Form,
        
        [Parameter(Mandatory=$true)]
        [string]$Status
    )
    $Label.Text = "Status: $Status..."
    $Form.Refresh()
}
