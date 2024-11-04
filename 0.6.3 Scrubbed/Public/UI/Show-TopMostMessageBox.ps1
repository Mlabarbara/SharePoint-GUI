using namespace System.Windows.Forms
using namespace System.Drawing

function Show-TopMostMessageBox {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Message,
        
        [Parameter(Mandatory=$true, Position=1)]
        [string]$Title,
        
        [Parameter(Position=2)]
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        
        [Parameter(Position=3)]
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )

    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true

    return [System.Windows.Forms.MessageBox]::Show($form, $Message, $Title, $Buttons, $Icon)
}
