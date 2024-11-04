using namespace System.Windows.Forms
using namespace System.Drawing

function New-Label($Text, $X, $Y, $Width, $Height) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    $label.Anchor = 'Left, Top, Right'
    return $label
}
