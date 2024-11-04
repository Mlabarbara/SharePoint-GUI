using namespace System.Windows.Forms
using namespace System.Drawing

# Import required functions
# . (Join-Path $PSScriptRoot '..\Public\UI\New-StandardForm.ps1')
# . (Join-Path $PSScriptRoot '..\Public\UI\New-Label.ps1')
# . (Join-Path $PSScriptRoot '..\Public\UI\New-TextBox.ps1')
# . (Join-Path $PSScriptRoot '..\Public\UI\New-Button.ps1')
# . (Join-Path $PSScriptRoot '..\Public\UI\New-ComboBox.ps1')
# . (Join-Path $PSScriptRoot '..\Private\Utility\Write-Log.ps1')

# end Show-CreateSubsiteForm
function Show-ConvertFolderToSubsite {
    [CmdletBinding()]
    param()

    try {
        Write-Log "Starting the ConvertFolderToSubsite Function"
        Show-StatusWindow "Getting Client folders from cache"
        $cachedData = Get-ClientJsonCache -selectedTeam "TeamSales"
        # grab all the not null clients from the cache
        $allFolders = @()
        if ($cachedData.General) {
            $allFolders += $cachedData.General
        }
        foreach ($team in @("TeamJamie", "TeamAmanda", "TeamKyle")) {
            if ($cachedData.ResearchClients.$team) {
                $allFolders += $cachedData.ResearchClients.$team
            }
        }
        Show-StatusWindow -Close
        
        # Create Subsite button
        if ($env:USERNAME -eq "marklabarbara") {
            $form = New-StandardForm 'Create Client Subsite' 420 400
            $skipEmailsCheckbox = New-Object System.Windows.Forms.CheckBox
            $skipEmailsCheckbox.Text = "Skip sending emails"
            $skipEmailsCheckbox.Location = New-Object System.Drawing.Point(10, 290)
            $skipEmailsCheckbox.Size = New-Object System.Drawing.Size(280, 20)
            $form.Controls.Add($skipEmailsCheckbox)
    
            $skipAlertsCheckbox = New-Object System.Windows.Forms.CheckBox
            $skipAlertsCheckbox.Text = "Skip creating alerts"
            $skipAlertsCheckbox.Location = New-Object System.Drawing.Point(10, 320)
            $skipAlertsCheckbox.Size = New-Object System.Drawing.Size(300, 20)
            $form.Controls.Add($skipAlertsCheckbox)
        } else {
            # If it's not marklababara, create hidden checkboxes set to false
            Write-Log "creating the main form"
            $form = New-StandardForm -title "Create Subsite from Folder" -width 420 -height 360
            $skipEmailsCheckbox = New-Object System.Windows.Forms.CheckBox
            $skipEmailsCheckbox.Checked = $false
            $skipEmailsCheckbox.Visible = $false
            $form.Controls.Add($skipEmailsCheckbox)
    
            $skipAlertsCheckbox = New-Object System.Windows.Forms.CheckBox
            $skipAlertsCheckbox.Checked = $false
            $skipAlertsCheckbox.Visible = $false
            $form.Controls.Add($skipAlertsCheckbox)
        }  # Increased height

        # Folder selection
        $folderLabel = New-Label -Text "Select Client Folder*:" -X 20 -Y 40 -Width 150 -Height 30  # Y increased by 20
        $folderDropdown = New-ComboBox -X 180 -Y 40 -Width 200 -Height 30  # Y increased by 20
        $form.Controls.AddRange(@($folderLabel, $folderDropdown))

        # Add note about client order
        $orderNote = New-Label -Text "*All clients are listed: General, Research TeamKyle, Research TeamJamie, Research TeamAmanda, and then alphabetically. " -X 20 -Y 10 -Width 380 -Height 30
        $orderNote.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
        $form.Controls.Add($orderNote)
        Write-Log "Populating folder dropdown"
        # Populate folder dropdown

        if ($allFolders -and $allFolders.Count -gt 0) {
            $displayNames = $allFolders | ForEach-Object {
                if ($_.Location -eq "General") {
                    "$($_.ClientName) (General)"
                } else {
                    "$($_.ClientName) ($($_.Url -match 'Research Clients/(Team\w+)/' ? $matches[1] : 'Research'))"
                }
            }
            $folderDropdown.Items.AddRange($displayNames)
        }
        Write-Log "Folder dropdown populated"

        # Team selection (initially hidden)
        $teamLabel = New-Label -Text "Select Team:" -X 20 -Y 80 -Width 150 -Height 30
        $teamDropdown = New-ComboBox -X 180 -Y 80 -Width 200 -Height 30 -Items @("TeamJamie", "TeamAmanda", "TeamKyle")
        $teamLabel.Visible = $false
        $teamDropdown.Visible = $false
        $form.Controls.AddRange(@($teamLabel, $teamDropdown))

        # Email field
        $emailLabel = New-Label -Text "Client Email:" -X 20 -Y 100 -Width 150 -Height 30
        $emailTextBox = New-TextBox -X 180 -Y 120 -Width 200 -Height 30
        $emailLabel.Visible = $false
        $emailTextBox.Visible = $false
        $form.Controls.AddRange(@($emailLabel, $emailTextBox))

        # Email display
        $emailDisplayLabel = New-Label -Text "Shared Email:" -X 20 -Y 120 -Width 150 -Height 30
        $emailDisplayTextBox = New-TextBox -X 180 -Y 120 -Width 200 -Height 30
        $emailDisplayTextBox.ReadOnly = $true
        $form.Controls.AddRange(@($emailDisplayLabel, $emailDisplayTextBox))

        # Add "Send Email to Payments" checkbox
        $paymentsCheckbox = New-Object System.Windows.Forms.CheckBox
        $paymentsCheckbox.Text = "Send Email to Payments"
        $paymentsCheckbox.Location = New-Object System.Drawing.Point(20, 160)
        $paymentsCheckbox.Size = New-Object System.Drawing.Size(200, 30)
        $form.Controls.Add($paymentsCheckbox)

        # add the status label 
        $statusLabel = New-Label "status: Ready" 20 260 380 30
        $form.Controls.Add($statusLabel)

        $createButton = New-Button -Text "Create Subsite" -X 160 -Y 220 -Width 100 -Height 30 -OnClick {
            Show-StatusWindow "Sending information to New-ClientSubsite"
            $skipEmails = $skipEmailsCheckbox.Checked
            $skipAlerts = $skipAlertsCheckbox.Checked
            try {
                $selectedFolder = $allFolders | Where-Object { 
                    $dropdownText = $folderDropdown.SelectedItem
                    if ($dropdownText -match '^(.*?)\s+\((General|Team\w+)\)$') {
                        $_.ClientName -eq $matches[1] -and (
                            ($matches[2] -eq "General" -and $_.Location -eq "General") -or
                            ($matches[2] -ne "General" -and $_.Url -match [regex]::Escape($matches[2]))
                        )
                    }
                }
                if (-not $selectedFolder) {
                    [System.Windows.Forms.MessageBox]::Show("Please select a folder.")
                    return
                }
                $team = $script:selectedTeam ?? $teamDropdown.SelectedItem
                if (-not $team) {
                    [System.Windows.Forms.MessageBox]::Show("You must select a Team.")
                    return "Error"
                }
                $clientName = $selectedFolder.ClientName
                $clientEmail = if ($emailDisplayTextBox.Text) { $emailDisplayTextBox.Text } else { $emailTextBox.Text }
                if ([string]::IsNullOrWhiteSpace($clientEmail)) {
                    [System.Windows.Forms.MessageBox]::Show("Please enter a client email address.")
                    return
                }
                $confirmResult = Show-TrickyMessageBox -ClientName $clientName -ClientEmail $clientEmail -TeamName $team -Title "Confirm Subsite Creation Details"
                if (-not $confirmResult) {
                    Write-Log "Client Information validation failed or was canceled"
                    Update-StatusLabel -Label $statusLabel -Form $form -Status "Operation cancelled. Please try again."
                    return
                }
                Write-Log "Starting New-ClientSubsite function..."
                Write-Log "Sending... $clientName"
                Write-Log "Sending.. $clientEmail"
                Write-Log "Sending...$team"
                Write-Log "Sending...$skipEmails"
                Write-Log "Sending...$skipAlerts"
                Show-StatusWindow -Close
                #Create a job for New-ClientSubsite
                $newSubsiteJob = Start-Job -ScriptBlock {
                    param($clientName, $clientEmail, $team, $paymentsChecked, $skipEmails, $skipAlerts)
                    
                    # Import necessary modules
                    Import-Module "$script:LOCAL_ROOT\SharePointModule.psm1" -Force
                    Import-Module "$script:LOCAL_ROOT\ClientFunctions.ps1" -Force
                    Import-Module "$script:LOCAL_ROOT\HelperFunctions.psm1" -Force
                    Import-Module PnP.PowerShell

                    # Execute New-ClientSubsite
                    New-ClientSubsite -clientName $clientName -clientEmail $clientEmail -TeamName $team `
                    -skipEmails:$skipEmails -skipAlerts:$skipAlerts -sendPaymentsEmail:$paymentsChecked -EmailTemplateName "ConvertedClient"
                
                } -ArgumentList $clientName, $clientEmail, $team, $paymentsCheckbox.Checked, $skipEmails, $skipAlerts
                Write-Log "New-ClientSubsite job started moving"
                # Wait for the job to complete and get the result
                $newSubsiteResult = Receive-Job -Job $newSubsiteJob -Wait
                Remove-Job -Job $newSubsiteJob
                Write-Log "New-ClientSubsite job completed with result: $newSubsiteResult"

                if ($newSubsiteResult -like "Success") {
                    Show-StatusWindow "Subsite created successfully,  Removing Client from folder cache..."
                    # Create a job for Move-ClientDocuments
                    try {
                        #remove client from folder cache
                        Remove-ClientJsonCacheEntry -clientName $clientName -WhereIs "Folder"
                        Write-Log "Client: $clientName removed from folder cache"
                    }
                    catch {
                        Write-Log "Error removing client from folder cache: $($_.Exception.Message)" -Level "ERROR"
                    }
                    
                    $moveDocsJob = Start-Job -ScriptBlock {
                        param($clientName, $team, $selectedClientUrl)
                        
                        # Import necessary modules
                        Import-Module "$script:LOCAL_ROOT\SharePointModule.psm1" -Force
                        Import-Module "$script:LOCAL_ROOT\ClientFunctions.ps1" -Force
                        Import-Module "$script:LOCAL_ROOT\HelperFunctions.psm1" -Force
                        Import-Module PnP.PowerShell

                        # Execute Move-ClientDocuments
                        Move-ClientDocuments -clientName $clientName -TeamName $team -selectedClientUrl $selectedClientUrl
                    } -ArgumentList $clientName, $team, $selectedClientUrl
                    Write-Log "Move-ClientDocuments job started..."
                    Show-StatusWindow -Close
                    # Wait for the job to complete and get the result
                    $moveDocsResult = Receive-Job -Job $moveDocsJob -Wait
                    Remove-Job -Job $moveDocsJob
                    Write-Log "Move-ClientDocuments job completed with result: $($moveDocsResult | ConvertTo-Json -Compress)"

                    if ($moveDocsResult.Status -eq "Success") {
                        $messageText = "Subsite created successfully. "
                        $messageText += "Moved $($moveDocsResult.MovedFiles) files and $($moveDocsResult.MovedFolders) folders. "
                        
                        if ($moveDocsResult.FolderRemovalStatus -eq "Success") {
                            $messageText += "Original folder removed successfully."
                        } else {
                            $messageText += "Warning: Could not remove original folder. $($moveDocsResult.FolderRemovalMessage)"
                        }

                        [System.Windows.Forms.MessageBox]::Show($messageText, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        Write-Log $messageText
                        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    } elseif ($moveDocsResult.Status -eq "EmptyFolder") {
                        $messageText = "Subsite created successfully. The original folder was empty and has been removed."
                        [System.Windows.Forms.MessageBox]::Show($messageText, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        Write-Log $messageText
                        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    } else {
                        $errorMessage = "Error moving documents: $($moveDocsResult.Message)"
                        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        Write-Log $errorMessage -Level "ERROR"
                        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Error creating subsite: $newSubsiteResult", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    Write-Log "Error creating subsite: $newSubsiteResult" -Level "ERROR"
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show($form, "An error occurred: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
                $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            } finally {
                # Close the form 
                $form.Close()
            }
        }
        $form.Controls.Add($createButton)
        # add script level variable to store the Url
        $script:selectedClientUrl = $null
        # add a script level variable to store the selected team
        $script:selectedTeam = $null
        # evnt handler for team selection
        $teamDropdown.Add_SelectedIndexChanged({
            $script:selectedTeam = $teamDropdown.SelectedItem
            Write-Log "Team selected: $script:selectedTeam"
        })
        # Event handler for folder selection
        $folderDropdown.Add_SelectedIndexChanged({
            $selectedFolder = $allFolders | Where-Object { 
                $dropdownText = $folderDropdown.SelectedItem
                if ($dropdownText -match '^(.*?)\s+\((General|Team\w+)\)$') {
                    $_.ClientName -eq $matches[1] -and (
                        ($matches[2] -eq "General" -and $_.Location -eq "General") -or
                        ($matches[2] -ne "General" -and $_.Url -match [regex]::Escape($matches[2]))
                    )
                }
            }
            if ($selectedFolder) {
                Write-Log "Selected folder: $($selectedFolder.ClientName)"
                Write-Log "Selected folder Location: $($selectedFolder.Location)"
                Write-Log "Selected folder ClientEmails: $($selectedFolder.ClientEmails | ConvertTo-Json -Compress)"
                
                if ($selectedFolder.Location -eq "General") {
                    $teamLabel.Visible = $true
                    $teamDropdown.Visible = $true
                    $script:selectedTeam = $null
                } else {
                    $teamLabel.Visible = $false
                    $teamDropdown.Visible = $false
                    # search the $selectedFolder.Url for the team name and assign that to $script:selectedTeam
                    if ($selectedFolder.Url -match 'Research Clients/(Team\w+)/') {
                        $script:selectedTeam = $matches[1]
                        Write-Log "Team determined from folder URL: $script:selectedTeam"
                    } else {
                        Write-Log "Error: Could not determine team from folder URL: $($selectedFolder.Url)" -Level "ERROR"
                        $script:selectedTeam = $null
                    }
                }
                
                $script:selectedClientUrl = $selectedFolder.Url
                Write-Log "Selected Client Url: $script:selectedClientUrl"
                
                if($selectedFolder.ClientEmails -and $selectedFolder.ClientEmails.Count -gt 0){
                    $emailToDisplay = $selectedFolder.ClientEmails[0]
                    Write-Log "Email to display: $emailToDisplay"
                    $emailDisplayTextBox.Text = $emailToDisplay
                    $emailDisplayLabel.Visible = $true
                    $emailDisplayTextBox.Visible = $true
                    $emailLabel.Visible = $false
                    $emailTextBox.Visible = $false
                    Write-Log "Set email display text box to: $($emailDisplayTextBox.Text)"
                } else {
                    $emailDisplayTextBox.Text = ""
                    $emailDisplayLabel.Visible = $false
                    $emailDisplayTextBox.Visible = $false
                    $emailLabel.Visible = $true
                    $emailTextBox.Visible = $true
                    Write-Log "No email found for client, cleared email display text box"
                }
                Update-StatusLabel -Label $statusLabel -Form $form -Status "Ready"
            } else {
                Write-Log "No folder selected"
            }
        })
        $cleanupPerformed = $false
        $form.Add_FormClosing({
            param($sender1, $e)
            Write-Log "FormClosing event triggered"
            if (-not $cleanupPerformed) {
                Write-Log "Performing FormClose Cleanup..."
                Invoke-Cleanup -Form $form
                $script:cleanupPerformed = $true
                Write-Log "CreateSubsite cleaned up exiting..."
            }
        })

        Write-Log "Displaying ConvertFolderToSubsite for $env:USERNAME"
        $dialogResult = $form.ShowDialog()
        
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "ConvertFolderToSubsite completed successfully"
            return "Success"
        } else {
            Write-Log "ConvertFolderToSubsite was cancelled or encountered an error"
            return $form.DialogResult
        }
    }
    catch {
        Write-Log "Error in ConvertFolderToSubsite: $_" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return "Error"
    }
    finally {
        Write-Log "$env:USERNAME closed ConvertFolderToSubsite"
        if ($form -and -not $form.IsDisposed) {
            $form.Dispose()
        }
    }
}
