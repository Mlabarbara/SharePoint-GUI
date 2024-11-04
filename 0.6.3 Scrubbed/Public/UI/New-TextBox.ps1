using namespace System.Windows.Forms
using namespace System.Drawing

function New-TextBox($X, $Y, $Width, $Height) {
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point($X, $Y)
    $textBox.MinimumSize = New-Object System.Drawing.Size($Width, $Height)
    $textBox.Anchor = 'Top, Left, Right'
    $textBox.AutoSize = $true
    return $textBox
}
