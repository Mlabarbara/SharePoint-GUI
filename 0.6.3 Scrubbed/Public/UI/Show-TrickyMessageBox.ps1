using namespace System.Windows.Forms
using namespace System.Drawing

function Show-TrickyMessageBox {
    param (
        [string]$ClientName,
        [string]$ClientEmail,
        [string]$Title,
        [string]$TeamName = $null
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(400,300)  # Increased height to accommodate custom title bar
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.BackColor = [System.Drawing.Color]::White

    # Custom title bar
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.Dock = [System.Windows.Forms.DockStyle]::Top
    $titleBar.Height = 30
    $titleBar.BackColor = [System.Drawing.Color]::FromArgb(135, 206, 250)  # Light Sky Blue

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.ForeColor = [System.Drawing.Color]::Black
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.AutoSize = $true
    $titleLabel.Location = New-Object System.Drawing.Point(5, 5)
    $titleBar.Controls.Add($titleLabel)

    $form.Controls.Add($titleBar)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(15,50)  # Moved down to accommodate title bar
    $label.Size = New-Object System.Drawing.Size(380,200)
    $label.Text = "Are you sure you want to send emails to the following recipients?`n`n"
    if ($TeamName) {
        $label.Text += "Team Name:      $TeamName`n"
    }
    $label.Text += "Client:                 $ClientName`n"

    $emails = $ClientEmail -split '\s*;\s*' | Where-Object { $_ -ne '' }
    if ($emails.Count -gt 0) {
        $label.Text += "Recipients:`n"
        $label.Text += ($emails | ForEach-Object { "                         $_" }) -join "`n"
    } else {
        $label.Text += "Recipients:         (No recipients)`n"
    }
    $form.Controls.Add($label)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(135,250)  # Moved down
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(225,250)  # Moved down
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $form.Controls.Add($cancelButton)
    
    $moveCount = 0
    $mouseEnterHandler = {
        if ($moveCount -eq 0) {
            $random = New-Object System.Random
            $xOffset = $random.Next(-150, 150)
            $yOffset = $random.Next(-150, 150)
            $newX = $form.Location.X + $xOffset
            $newY = $form.Location.Y + $yOffset
            $form.Location = New-Object System.Drawing.Point($newX, $newY)
            $moveCount = 1
            $okButton.Remove_MouseEnter($mouseEnterHandler)
        }
    }
    $okButton.Add_MouseEnter($mouseEnterHandler)

    $okButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    $result = $form.ShowDialog()
    return $result -eq [System.Windows.Forms.DialogResult]::OK
}
