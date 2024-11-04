using namespace System.Windows.Forms
using namespace System.Drawing

# Import required functions
. (Join-Path $PSScriptRoot '..\Public\UI\New-StandardForm.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-Label.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-TextBox.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-Button.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-ComboBox.ps1')
. (Join-Path $PSScriptRoot '..\Private\Utility\Write-Log.ps1')

# End Show-ConvertGeneralToResearchForm
function Show-CreateSubsiteForm {
    [CmdletBinding()]
    param()

    # Create the form with increased height to accommodate new controls
    # $subSiteForm = New-StandardForm 'Create Client Subsite' 320 340
    # New checkboxes under the button to send $skipEmails and $skipAlerts to New-ClientSubsite
    if ($env:USERNAME -eq "marklabarbara") {
        $subSiteForm = New-StandardForm 'Create Client Subsite' 320 400
        $skipEmailsCheckbox = New-Object System.Windows.Forms.CheckBox
        $skipEmailsCheckbox.Text = "Skip sending emails"
        $skipEmailsCheckbox.Location = New-Object System.Drawing.Point(10, 290)
        $skipEmailsCheckbox.Size = New-Object System.Drawing.Size(260, 20)
        $subSiteForm.Controls.Add($skipEmailsCheckbox)

        $skipAlertsCheckbox = New-Object System.Windows.Forms.CheckBox
        $skipAlertsCheckbox.Text = "Skip creating alerts"
        $skipAlertsCheckbox.Location = New-Object System.Drawing.Point(10, 320)
        $skipAlertsCheckbox.Size = New-Object System.Drawing.Size(260, 20)
        $subSiteForm.Controls.Add($skipAlertsCheckbox)
    } else {
        # If it's not marklababara, create hidden checkboxes set to false
        $subSiteForm = New-StandardForm 'Create Client Subsite' 320 340
        $skipEmailsCheckbox = New-Object System.Windows.Forms.CheckBox
        $skipEmailsCheckbox.Checked = $false
        $skipEmailsCheckbox.Visible = $false
        $subSiteForm.Controls.Add($skipEmailsCheckbox)

        $skipAlertsCheckbox = New-Object System.Windows.Forms.CheckBox
        $skipAlertsCheckbox.Checked = $false
        $skipAlertsCheckbox.Visible = $false
        $subSiteForm.Controls.Add($skipAlertsCheckbox)
    }

    # Existing Controls
    $subSiteForm.Controls.Add((New-Label 'Client Name:' 10 20 100 20))
    $clientNameTextBox = New-TextBox 120 20 150 20
    $subSiteForm.Controls.Add($clientNameTextBox)
    
    $subSiteForm.Controls.Add((New-Label 'Client Email:' 10 50 100 20))
    $clientEmailTextBox = New-TextBox 120 50 150 20
    $subSiteForm.Controls.Add($clientEmailTextBox)
    
    $subSiteForm.Controls.Add((New-Label 'Select Team:' 10 80 100 20))
    $teamComboBox = New-ComboBox 120 80 150 20 @('TeamJamie', 'TeamAmanda', 'TeamKyle')
    $subSiteForm.Controls.Add($teamComboBox)
    
    $sendToPaymentsCheckbox = New-Object System.Windows.Forms.CheckBox
    $sendToPaymentsCheckbox.Text = "Send Email to Payments"
    $sendToPaymentsCheckbox.Location = New-Object System.Drawing.Point(10, 110)
    $sendToPaymentsCheckbox.Size = New-Object System.Drawing.Size(260, 20)
    $subSiteForm.Controls.Add($sendToPaymentsCheckbox)
    
    # New Checkbox for selecting user
    $selectUserCheckbox = New-Object System.Windows.Forms.CheckBox
    $selectUserCheckbox.Text = "Would you like to select the user from which to send the email?"
    $selectUserCheckbox.Location = New-Object System.Drawing.Point(10, 140)
    $selectUserCheckbox.Size = New-Object System.Drawing.Size(260, 40)
    $subSiteForm.Controls.Add($selectUserCheckbox)
    
    # New Combobox for user selection, initially hidden
    $userComboBox = New-ComboBox
    $userComboBox.Items.AddRange(@('Jeff', 'Fred', 'Levi', 'Jamie', 'AmandaS', 'Naomi', 'DonnaRae', 'Crystal', 'Brande', 'Dakota'))
    $userComboBox.Location = New-Object System.Drawing.Point(10, 190)
    $userComboBox.Size = New-Object System.Drawing.Size(260, 20)
    $userComboBox.Visible = $false
    $subSiteForm.Controls.Add($userComboBox)
    
    # Event Handler to toggle combobox visibility
    $selectUserCheckbox.Add_CheckedChanged({
        if ($selectUserCheckbox.Checked) {
            $userComboBox.Visible = $true
        } else {
            $userComboBox.Visible = $false
            $userComboBox.SelectedIndex = -1
        }
    })
    
    # Status Label
    $statusLabel = New-Label "Status: Ready" 10 220 280 20
    $subSiteForm.Controls.Add($statusLabel)
    
    # Create Subsite Button
    $createSubsiteButton = New-Button 'Create Subsite' 75 250 150 30 {
        # Retrieve form values
        $clientName = $clientNameTextBox.Text
        $clientEmail = $clientEmailTextBox.Text
        $selectedTeam = $teamComboBox.SelectedItem
        $sendToPayments = $sendToPaymentsCheckbox.Checked
        $selectUser = $selectUserCheckbox.Checked
        $selectedUser = if ($selectUser) { $userComboBox.SelectedItem } else { $null }
        $skipEmails = $skipEmailsCheckbox.Checked
        $skipAlerts = $skipAlertsCheckbox.Checked
        
        # Validate required fields
        if ([string]::IsNullOrWhiteSpace($clientName) -or 
            [string]::IsNullOrWhiteSpace($clientEmail) -or 
            [string]::IsNullOrWhiteSpace($selectedTeam)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all required fields.", "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Validate user selection if checkbox is checked
        if ($selectUser -and [string]::IsNullOrWhiteSpace($selectedUser)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a user to send the email from.", "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Confirm Details
        $confirmResult = Show-TrickyMessageBox -clientName $clientName -clientEmail $clientEmail -TeamName $selectedTeam -Title "Confirm Details for Subsite Creation" 
        if (-not $confirmResult) {
            Write-Log "Client Information validation failed or was canceled"
            Update-StatusLabel -Label $statusLabel -Form $subSiteForm -Status "Operation cancelled. Please try again."
            return
        }
        
        # Update Status
        Update-StatusLabel -Label $statusLabel -Form $subSiteForm -Status "Creating subsite..."
        
        try {
            Write-Log "Starting New-ClientSubsite function in a new runspace..."
            Write-Log "Sending... $clientName"
            Write-Log "Sending.. $clientEmail"
            Write-Log "Sending...$selectedTeam"
            Write-Log "Sending...$sendToPayments"
            Write-Log "Sending...$selectUser...$selectedUser"
            Write-Log "Sending...$skipEmails"
            Write-Log "Sending...$skipAlerts"
    
            # Determine the From email address
            if ($selectUser) {
                $fromEmail = "$selectedUser@sequoiataxrelief.com"
                Write-Log "Using selected user email as From: $fromEmail"
            } else {
                $fromEmail = "$env:USERNAME@sequoiataxrelief.com"
                Write-Log "Using current user email as From: $fromEmail"
            }
    
            # Create and open a PowerShell runspace
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()
            $runspace.SessionStateProxy.SetVariable("clientName", $clientName)
            $runspace.SessionStateProxy.SetVariable("clientEmail", $clientEmail)
            $runspace.SessionStateProxy.SetVariable("team", $selectedTeam)
            $runspace.SessionStateProxy.SetVariable("sendPaymentsEmail", $sendToPayments)
            $runspace.SessionStateProxy.SetVariable("fromEmail", $fromEmail)
            $runspace.SessionStateProxy.SetVariable("skipEmails", $skipEmails)
            $runspace.SessionStateProxy.SetVariable("skipAlerts", $skipAlerts)
            
    
            # Create a PowerShell instance and add the script
            $powershell = [powershell]::Create().AddScript({
                param($clientName, $clientEmail, $team, $sendPaymentsEmail, $fromEmail, $skipEmails, $skipAlerts)
    
                # Import necessary modules
                Import-Module "$script:LOCAL_ROOT\SharePointModule.psm1" -Force
                Import-Module "$script:LOCAL_ROOT\ClientFunctions.ps1" -Force
                Import-Module "$script:LOCAL_ROOT\HelperFunctions.psm1" -Force
                Import-Module PnP.PowerShell
    
                $result = New-ClientSubsite -clientName $clientName -clientEmail $clientEmail -TeamName $team `
                        -sendPaymentsEmail:$sendPaymentsEmail -From $fromEmail -skipEmails:$skipEmails -skipAlerts:$skipAlerts -EmailTemplateName "GeneralClient"

                return $result
            }).AddArgument($clientName).AddArgument($clientEmail).AddArgument($selectedTeam).AddArgument($sendToPayments).AddArgument($fromEmail).AddArgument($skipEmails).AddArgument($skipAlerts)
    
            # Associate the runspace with the PowerShell instance and run it asynchronously
            $powershell.Runspace = $runspace
            $asyncResult = $powershell.BeginInvoke()
    
            # Wait for the operation to complete with a timeout
            $completed = $asyncResult.AsyncWaitHandle.WaitOne(300000) # 5-minute timeout
    
            if ($completed) {
                $result = $powershell.EndInvoke($asyncResult)
                Write-Log "New-ClientSubsite function completed with result: $result"
    
                if ($result -eq "Success") {
                    Update-StatusLabel -Label $statusLabel -Form $subSiteForm -Status "Subsite created successfully"
                    Write-Log "Subsite created successfully"
                    $subSiteForm.DialogResult = [System.Windows.Forms.MessageBox]::Show("Subsite created successfully and emails sent.", "Success", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                } else {
                    Update-StatusLabel -Label $statusLabel -Form $subSiteForm -Status "Error: $result"
                    Write-Log "New-ClientSubsite function completed with errors: $result" -Level "ERROR"
                    $subSiteForm.DialogResult = [System.Windows.Forms.MessageBox]::Show("Error: $result", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                Write-Log "New-ClientSubsite function timed out after 5 minutes" -Level "ERROR"
                Update-StatusLabel -Label $statusLabel -Form $subSiteForm -Status "Operation timed out"
                $subSiteForm.DialogResult = [System.Windows.Forms.MessageBox]::Show("The operation timed out after 5 minutes. Please check the logs for more information.", 
                    "Timeout Error", [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } catch {
            Write-Log "An error occurred: $_" -Level "ERROR"
            Update-StatusLabel -Label $statusLabel -Form $subSiteForm -Status "Error occurred"
            $subSiteForm.DialogResult = [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            if ($powershell) {
                $powershell.Dispose()
            }
            if ($runspace) {
                $runspace.Dispose()
            }
            Update-StatusLabel -Label $statusLabel -Form $subSiteForm -Status "Operation completed"
        }
    }
    $subSiteForm.Controls.Add($createSubsiteButton)

    try {
        Write-Log "Displaying CreateSubsite for $env:USERNAME"
        $dialogResult = $subSiteForm.ShowDialog()
    
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "CreateSubsite completed successfully"
            return "Success"
        } else {
            Write-Log "CreateSubsite was cancelled or closed"
            return 2
        }
    } 
    finally {
        Write-Log "$env:USERNAME closed CreateSubsite"
        $subSiteForm.Dispose()
    }
}
