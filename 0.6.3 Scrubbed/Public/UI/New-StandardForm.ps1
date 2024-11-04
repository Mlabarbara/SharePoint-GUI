using namespace System.Windows.Forms
using namespace System.Drawing

function New-StandardForm($title, $width, $height) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Width = $width
    $form.Height = $height
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'Sizable'
    $form.MinimumSize = New-Object System.Drawing.Size($width, $height)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $form.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $form.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
    $form.AutoScroll = $true
    $form.MaximizeBox = $false
    $form.BringToFront()
    return $form
}
