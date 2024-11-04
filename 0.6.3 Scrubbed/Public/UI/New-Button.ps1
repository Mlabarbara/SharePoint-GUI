using namespace System.Windows.Forms
using namespace System.Drawing

function New-Button($Text, $X, $Y, $Width, $Height, $OnClick) {
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    $button.Add_Click($OnClick)
    $button.Anchor = 'Top, Left, Right'
    return $button
}
