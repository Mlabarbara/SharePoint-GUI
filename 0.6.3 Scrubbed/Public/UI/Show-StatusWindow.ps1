using namespace System.Windows.Forms
using namespace System.Drawing

## I think this will replace the Show-LoadingForm function
function Show-StatusWindow {
    param (
        [string]$Status = "Starting...",
        [switch]$Close
    )
    
    if (-not $script:statusForm) {
        $script:statusForm = New-Object System.Windows.Forms.Form
        $script:statusForm.Text = "Operation Status"
        $script:statusForm.Size = New-Object System.Drawing.Size(300,150)
        $script:statusForm.StartPosition = 'CenterScreen'
        $script:statusForm.FormBorderStyle = 'FixedDialog'
        $script:statusForm.MaximizeBox = $false
        $script:statusForm.MinimizeBox = $false
        $script:statusForm.Topmost = $true

        $script:statusLabel = New-Object System.Windows.Forms.Label
        $script:statusLabel.Location = New-Object System.Drawing.Point(10,20)
        $script:statusLabel.Size = New-Object System.Drawing.Size(280,60)
        $script:statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $script:statusForm.Controls.Add($script:statusLabel)
    }

    if ($Close) {
        if ($script:statusForm -and $script:statusForm.Visible) {
            $script:statusForm.Invoke([Action]{$script:statusForm.Close()})
            $script:statusForm.Dispose()
            $script:statusForm = $null
        }
    } else {
        $script:statusLabel.Text = $Status
        if (-not $script:statusForm.Visible) {
            $script:statusForm.Show()
        }
        $script:statusForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
}
