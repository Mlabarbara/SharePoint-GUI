using namespace System.Windows.Forms
using namespace System.Drawing

function New-ComboBox($X, $Y, $Width, $Height, $Items) {
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point($X, $Y)
    $comboBox.Size = New-Object System.Drawing.Size($Width, $Height)
    $comboBox.Anchor = 'Top, Left, Right'
    if ($Items -and $Items.Count -gt 0) {
        $comboBox.Items.AddRange($Items)
    }
    return $comboBox
}
