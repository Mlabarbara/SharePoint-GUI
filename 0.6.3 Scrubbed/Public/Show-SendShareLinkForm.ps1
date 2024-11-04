using namespace System.Windows.Forms
using namespace System.Drawing

# Import required functions
. (Join-Path $PSScriptRoot '..\Public\UI\New-StandardForm.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-Label.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-TextBox.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-Button.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-ComboBox.ps1')
. (Join-Path $PSScriptRoot '..\Private\Utility\Write-Log.ps1')

# End ConvertFolderToSubsite 
function Show-SendShareLinkForm {
    [CmdletBinding()]
    param()

    $siteUrlBase = Get-SharePointBaseUrl

    $sendShareForm = New-StandardForm 'Resend Client Share Link' 330 490  # Increased Importform height

    $sendShareForm.Controls.Add((New-Label 'Select Team:' 10 20 100 20))
    $teamComboBox = New-ComboBox 120 20 150 20 @('TeamJamie', 'TeamAmanda', 'TeamKyle', 'TeamSales')
    $sendShareForm.Controls.Add($teamComboBox)

    $sendShareForm.Controls.Add((New-Label 'Select Client:' 10 50 100 20))
    $clientComboBox = New-ComboBox 120 50 150 20 @()
    $sendShareForm.Controls.Add($clientComboBox)
    # open the ComboBox dropdown when the first letter is pressed 
    $clientComboBox.Add_KeyPress({
        param($sender1, $e)
        if ($e.KeyChar -match '[a-zA-Z]') {
            $clientComboBox.DroppedDown = $true
        }
    })

    # Create a FlowLayoutPanel to hold email entries
    $emailFlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $emailFlowPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $emailFlowPanel.WrapContents = $false
    $emailFlowPanel.AutoScroll = $true
    $emailFlowPanel.Location = New-Object System.Drawing.Point(10, 80)
    $emailFlowPanel.Size = New-Object System.Drawing.Size(300, 120)  # Increased size
    $sendShareForm.Controls.Add($emailFlowPanel)

    $sendToSelfCheckbox = New-Object System.Windows.Forms.CheckBox
    $sendToSelfCheckbox.Text = "Send Link To Yourself?"
    $sendToSelfCheckbox.Location = New-Object System.Drawing.Point(10, 240)
    $sendToSelfCheckbox.Size = New-Object System.Drawing.Size(260, 20)
    $sendShareForm.Controls.Add($sendToSelfCheckbox)

    $additionalEmailCheckbox = New-Object System.Windows.Forms.CheckBox
    $additionalEmailCheckbox.Text = "Add an additional email?"
    $additionalEmailCheckbox.Location = New-Object System.Drawing.Point(10, 270)
    $additionalEmailCheckbox.Size = New-Object System.Drawing.Size(260, 20)
    $sendShareForm.Controls.Add($additionalEmailCheckbox)

    $additionalEmailTextBox = New-TextBox 10 300 260 20
    $additionalEmailTextBox.Visible = $false
    $sendShareForm.Controls.Add($additionalEmailTextBox)

    # Add note about client order
    $addEmailNote = New-Label -Text "*Enter additional email(s) separated by a semicolon ;" -X 10 -Y 330 -Width 380 -Height 25
    $addEmailNote.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Italic)
    $addEmailNote.Visible = $false
    
    # $statusLabel = New-Label "Status: Ready" 10 360 300 30
    # $sendShareForm.Controls.Add($statusLabel)

    $resendButton = New-Button 'Send Email' 75 400 160 40 { }  # We'll add the click event later
    $sendShareForm.Controls.Add($resendButton)

    $additionalEmailCheckbox.Add_CheckedChanged({
        $additionalEmailTextBox.Visible = $additionalEmailCheckbox.Checked
        $addEmailNote.Visible = $additionalEmailCheckbox.Checked
        if ($additionalEmailCheckbox.Checked) {
            $additionalEmailTextBox.Text = ""
            # Use BeginInvoke to set focus after the UI has updated
            $additionalEmailTextBox.BeginInvoke([Action]{
                $additionalEmailTextBox.Focus()
            })
        }
    })
    $sendShareForm.Controls.Add($addEmailNote)
    # Function to add email entry to the panel
    function Add-EmailEntry {
        param (
            [string]$Email,
            [bool]$IsExisting = $true
        )
        $emailPanel = New-Object System.Windows.Forms.Panel
        $emailPanel.Size = New-Object System.Drawing.Size(280, 30)

        $emailTextBox = New-TextBox 0 5 200 20
        $emailTextBox.Text = $Email
        $emailTextBox.ReadOnly = $IsExisting
        $emailPanel.Controls.Add($emailTextBox)

        $resendCheckBox = New-Object System.Windows.Forms.CheckBox
        $resendCheckBox.Text = "Resend?"
        $resendCheckBox.Location = New-Object System.Drawing.Point(210, 5)
        $resendCheckBox.Size = New-Object System.Drawing.Size(70, 20)
        $emailPanel.Controls.Add($resendCheckBox)

        $emailFlowPanel.Controls.Add($emailPanel)
    }
    $script:currentClientInfo = $null
    $script:teamClientData = @{}
    function Add-EmailToSharingLink {
        param (
            [string]$sharingLink,
            [string]$email
        )
        if ([string]::IsNullOrWhiteSpace($sharingLink)) {
            return $sharingLink
        }
        $uri = [System.Uri]$sharingLink
        $uriBuilder = New-Object System.UriBuilder($uri)
        $query = [System.Web.HttpUtility]::ParseQueryString($uriBuilder.Query)
        if (-not $query["email"]) {
            $query["email"] = $email
            $uriBuilder.Query = $query.ToString()
        }
        return $uriBuilder.Uri.ToString()
    }
    function Reset-Form {
        if ($sendShareForm -and !$sendShareForm.IsDisposed) {
            if ($sendShareForm.InvokeRequired) {
                $sendShareForm.Invoke([Action]{
                    Write-Log "Reset-Form invoked"
                    $emailFlowPanel.Controls.Clear()
                    $clientComboBox.SelectedIndex = -1
                    $additionalEmailCheckbox.Checked = $false
                    $additionalEmailTextBox.Text = ""
                    $additionalEmailTextBox.Visible = $false
                    $sendToSelfCheckbox.Checked = $false
                    $script:currentClientInfo = $null
                })
            } else {
                Write-Log "Clearing email fields..."
                $emailFlowPanel.Controls.Clear()
                $clientComboBox.SelectedIndex = -1
                $additionalEmailCheckbox.Checked = $false
                $additionalEmailTextBox.Text = ""
                $additionalEmailTextBox.Visible = $false
                $sendToSelfCheckbox.Checked = $false
                $script:currentClientInfo = $null
            }
        }
    }
    $teamComboBox.Add_SelectedIndexChanged({
        $clientComboBox.Items.Clear()
        $clientComboBox.Text = ""
        $clientComboBox.SelectedIndex = -1
        Write-Log "Clearing email fields"
        Reset-Form
        
        $selectedTeam = $teamComboBox.SelectedItem 
        if (-not $selectedTeam) { return }
        
        try {
            Show-StatusWindow "Retrieving clients..."
            Write-Log "Retrieving data for $selectedTeam"
            
            # Get client list based on team selection
            $clientList = if ($selectedTeam -eq "TeamSales") {
                $clientData = Get-ClientJsonCache -selectedTeam "TeamSales" -selectedClient $null
                Write-Log "TeamSales data retrieved successfully"
                
                # Combine all teams' data into one list
                @(
                    $clientData.General,
                    $clientData.ResearchClients.TeamJamie,
                    $clientData.ResearchClients.TeamAmanda,
                    $clientData.ResearchClients.TeamKyle
                ) | Where-Object { $null -ne $_ }
            } else {
                Get-ClientJsonCache -selectedTeam $selectedTeam -selectedClient $null
            }
    
            Write-Log "Client list created with $($clientList.Count) items"
            
            if ($clientList -and $clientList.Count -gt 0) {
                $script:teamClientData = @{}
                $clientList | Where-Object { 
                    $null -ne $_ -and $null -ne $_.ClientName 
                } | ForEach-Object {
                    $clientComboBox.Items.Add($_.ClientName)
                    $script:teamClientData[$_.ClientName] = $_
                }
                Write-Log "Added $($clientComboBox.Items.Count) clients to combo box"
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "No client data found for the selected team.", 
                    "Information", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                Write-Log "No clients found for $selectedTeam"
            }
        }
        catch {
            Write-Log "Error retrieving clients: $($_.Exception.Message)" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Error retrieving clients: $($_.Exception.Message)", 
                "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            Reset-Form
        }
        finally {
            Show-StatusWindow -Close
        }
    })
    # Modify the client selection event
    $clientComboBox.Add_SelectedIndexChanged({
        $selectedTeam = $teamComboBox.SelectedItem
        $selectedClient = $clientComboBox.SelectedItem
        if ($selectedTeam -and $selectedClient) {
            try {
                #retriece client info from the in-memory hashtable
                if ($script:teamClientData.ContainsKey($selectedClient)) {
                    $clientInfo = $script:teamClientData[$selectedClient]
                }
                else {
                    $clientInfo = Get-ClientJsonCache -selectedTeam $selectedTeam -selectedClient $selectedClient
                }

                #$clientInfo = Get-ClientJsonCache -selectedTeam $selectedTeam -selectedClient $selectedClient
                # Clear existing email entries
                $emailFlowPanel.Controls.Clear()
                $script:currentClientInfo = $null
    
                if ($clientInfo) {
                    $existingEmails = $clientInfo.ClientEmails
                    $clientSharingLink = $clientInfo.SharingLink
                    Write-Log "Client information loaded: $existingEmails | $clientSharingLink"
    
                    if ($existingEmails -and $clientSharingLink) {
                        if ($existingEmails -is [string]) {
                            $existingEmails = @($existingEmails)
                        }
                        
                        foreach ($email in $existingEmails) {
                            Add-EmailEntry -Email $email -IsExisting $true
                        }
                        
                        $script:currentClientInfo = @{
                            Emails = $existingEmails
                            SharingLink = $clientSharingLink
                        }
                        
                        # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Client information loaded"
                    } else {
                        # No email or sharing link found
                        $createNewLink = [System.Windows.Forms.MessageBox]::Show(
                            "No existing share link or email found. Would you like to create one?",
                            "Create New Share Link",
                            [System.Windows.Forms.MessageBoxButtons]::YesNo,
                            [System.Windows.Forms.MessageBoxIcon]::Question
                        )
                        
                        if ($createNewLink -eq [System.Windows.Forms.DialogResult]::Yes) {
                            Write-Log "User selected to create a new share link"
                            $additionalEmailCheckbox.Checked = $true
                            $additionalEmailTextBox.Visible = $true
                            $additionalEmailTextBox.Focus()
                            # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Please enter an email address to create a new share link"
                        } else {
                            # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "No share link or email available"
                        }
                    }
                } else {
                    # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Client information not found"
                    [System.Windows.Forms.MessageBox]::Show("Client information not found in the JSON file.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            } catch {
                # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Error checking client information"
                [System.Windows.Forms.MessageBox]::Show("Error checking client information: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    # Modify the resend button click event
    $resendButton.Add_Click({
        $selectedTeam = $teamComboBox.SelectedItem
        $selectedClient = $clientComboBox.SelectedItem
        $sendToSelf = $sendToSelfCheckbox.Checked
        $addAdditionalEmail = $additionalEmailCheckbox.Checked
        $additionalEmailInput = ($additionalEmailTextBox.Text -split ';' | ForEach-Object { $_.Trim() }) -join ';'
    
        Write-Log "Sending Share Link to $selectedClient in $selectedTeam"
    
        # Log if additional email is being added
        if ($addAdditionalEmail) {
            Write-Log "AddAdditionalEmail = True | Sending Share Link to $selectedClient in $selectedTeam and also sending to additional email address: $additionalEmailInput"
        }
        
        # Validate team and client selection
        if ([string]::IsNullOrWhiteSpace($selectedTeam) -or [string]::IsNullOrWhiteSpace($selectedClient)) {
            [System.Windows.Forms.MessageBox]::Show("Please select both a team and a client.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    
        # Validate additional email input if checked
        if ($addAdditionalEmail -and [string]::IsNullOrWhiteSpace($additionalEmailInput)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter at least one additional email to send the link to.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    
        # Collect existing selected emails
        $existingSelectedEmails = @()
        foreach ($control in $emailFlowPanel.Controls) {
            $emailTextBox = $control.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] }
            $resendCheckBox = $control.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] }
            if ($resendCheckBox.Checked) {
                $emails = $emailTextBox.Text -split ";" | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                $existingSelectedEmails += $emails
            }
        }
    
        # Collect additional new emails
        $additionalSelectedEmails = @()
        if ($addAdditionalEmail -and -not [string]::IsNullOrWhiteSpace($additionalEmailInput)) {
            $additionalSelectedEmails = $additionalEmailInput -split ";" | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
    
        # Ensure no duplicates
        $existingSelectedEmails = $existingSelectedEmails | Select-Object -Unique
        $additionalSelectedEmails = $additionalSelectedEmails | Select-Object -Unique
    
        # Log collected emails before appending self email
        Write-Log "Existing emails: $existingSelectedEmails | Additional Emails: $additionalSelectedEmails | Send to self?: $sendToSelf"
    
        
        # Combine all emails for confirmation
        $allEmails = $existingSelectedEmails + $additionalSelectedEmails | Select-Object -Unique
        if($sendToSelf){
            $allEmails += "$env:USERNAME@sequoiataxrelief.com"
        }
        if ($allEmails.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one email to send the link to.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    
        # Prepare email list for confirmation
        $emailList = $allEmails -join ";"
        $confirmResult = Show-TrickyMessageBox -ClientName $selectedClient -ClientEmail $emailList -TeamName $selectedTeam -Title "Confirm Share Link Recipients"
        if (-not $confirmResult) {
            Write-Log "Client Information validation failed or was canceled"
            # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Operation cancelled by user"
            Reset-Form
            return
        }
    
        # Proceed with sending emails
        $resendButton.Enabled = $false
        # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Processing..."
        try {
            Show-StatusWindow "Verifying client information..."
            # Fetch client information
            $clientInfo = Get-ClientJsonCache -selectedTeam $selectedTeam -selectedClient $selectedClient
            if (-not $clientInfo) {
                throw "Client information not found."
            }
    
            # Determine site URL and folder path
            if ($selectedTeam -eq "TeamSales") {
                $siteUrl = "https://sequoiataxrelief.sharepoint.com/sites/secureclientupload"
                $folderPath = $clientInfo.Url -replace "^/sites/secureclientupload/", ""
            } else {
                $sanitizedClientName = Remove-SpacesAndPunc $selectedClient
                $siteUrl = "$siteUrlBase/$selectedTeam/$sanitizedClientName"
                $folderPath = "Documents/Client Documents"
            }
            Show-StatusWindow "Connecting to sharepoint..."
            # Connect to SharePoint site
Connect-PnPOnline
    
            # Handle sharing link
            if (-not $script:currentClientInfo -or -not $script:currentClientInfo.SharingLink -or $addAdditionalEmail) {
                if (-not $script:currentClientInfo -or -not $script:currentClientInfo.SharingLink) {
                    # Create a new sharing link
                    Show-StatusWindow "Creating new sharing link..."
                    $sharingResult = Add-PnPFolderUserSharingLink -Folder $folderPath -ShareType Edit -Users $additionalSelectedEmails
                    if ($sharingResult) {
                        $script:currentClientInfo = @{
                            Emails = $additionalSelectedEmails
                            SharingLink = $sharingResult.WebUrl
                        }
                        Write-Log "New sharing link created: $($script:currentClientInfo.SharingLink) for $($additionalSelectedEmails -join ', ')"
                        # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "New sharing link created."
                    } else {
                        throw "Failed to create sharing link."
                    }
                } elseif ($addAdditionalEmail -and -not [string]::IsNullOrWhiteSpace($additionalSelectedEmails)) {
                    # Add users to existing sharing link
                    Show-StatusWindow "Adding new emails to existing sharing link..."
                    $sharingResult = Add-PnPFolderUserSharingLink -Folder $folderPath -ShareType Edit -Users $additionalSelectedEmails
                    if ($sharingResult) {
                        Write-Log "Added $($additionalSelectedEmails -join ', ') to existing sharing link."
                        # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Users added to existing sharing link."
                    } else {
                        throw "Failed to add users to existing sharing link."
                    }
                }
    
                # Update the JSON cache with new emails
                $existingEmails = $clientInfo.ClientEmails
                if ($existingEmails -is [string]) {
                    $existingEmails = @($existingEmails)
                }
                $updatedEmails = $existingEmails + $additionalSelectedEmails | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                Show-StatusWindow "Updating client information..."
                if ($selectedTeam -eq "TeamSales") {
                    # Update client_folder_cache.json for TeamSales
                    $allClientData = Get-ClientJsonCache -selectedTeam "TeamSales" -selectedClient $null
                    $sections = @('General', 'ResearchClients.TeamAmanda', 'ResearchClients.TeamJamie')
                    foreach ($section in $sections) {
                        $clientToUpdate = Invoke-Expression "`$allClientData.$section | Where-Object { `$_.ClientName -eq '$selectedClient' }"
                        if ($clientToUpdate) {
                            $clientToUpdate.ClientEmails = $updatedEmails
                            $clientToUpdate.SharingLink = $script:currentClientInfo.SharingLink
                            break
                        }
                    }
                    $jsonPath = Get-CacheFilePath -CacheType "Folder"
                    $allClientData | ConvertTo-Json -Depth 10 | Set-Content $jsonPath
                } else {
                    # Update cache for other teams
                    Update-ClientJsonCache -teamName $selectedTeam -clientName $selectedClient -email $updatedEmails -sharingLink $script:currentClientInfo.SharingLink
                }
            } else {
                # Existing sharing link is available and no additional email is being added
                Write-Log "Using existing sharing link: $($script:currentClientInfo.SharingLink)"
                # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Using existing sharing link."
            }
            Show-StatusWindow "Prepairing to send emails"
            # Prepare placeholders for email templates
            $currentUserEmail = "$env:USERNAME@sequoiataxrelief.com"
            $placeholders = @{
                ClientName         = $selectedClient
                CompanyLogoUrl     = Get-CompanyUrl -UrlType "Logo"
                SupportSiteUrl     = Get-CompanyUrl -UrlType "Support"
                PrivacyPolicyUrl   = Get-CompanyUrl -UrlType "PrivacyPolicy"
                CurrentYear        = (Get-Date).Year
                TeamName           = $selectedTeam
            }
            function Use-Hash {
                param (
                    [hashtable]$Hashtable
                )
                $newHashtable = @{}
                foreach ($key in $Hashtable.Keys) {
                    $newHashtable[$key] = $Hashtable[$key]
                }
                return $newHashtable
            }
            
            $emailsToSend = @()
            Write-Log "Preparing to send emails to selected recipients."
            # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Preparing emails..."
    
            # Prepare emails for existing email addresses
            foreach ($recipientEmail in $existingSelectedEmails) {
                $emailPlaceholders = Use-Hash -Hashtable $placeholders
                $emailPlaceholders.ClientEmail = $recipientEmail
                $emailPlaceholders.ClientSharingLink = Add-EmailToSharingLink -sharingLink $script:currentClientInfo.SharingLink -email $recipientEmail
    
                $clientEmailTemplate = Get-EmailTemplate -TemplateName "ResendShareLink"
                $clientEmailBody = Format-EmailTemplate -Template $clientEmailTemplate -Placeholders $emailPlaceholders
    
                Write-Log "Preparing ResendShareLink email for: $recipientEmail"
                $emailsToSend += @{
                    From            = $currentUserEmail
                    To              = @($recipientEmail)
                    Subject         = "Requested Link to Your Secure Client Portal - Sequoia Tax Relief"
                    Body            = $clientEmailBody
                    BodyContentType = "Html"
                }
            }
    
            # Prepare emails for additional (new) email addresses
            foreach ($recipientEmail in $additionalSelectedEmails) {
                $emailPlaceholders = Use-Hash -Hashtable $placeholders
                $emailPlaceholders.ClientEmail = $recipientEmail
                $emailPlaceholders.ClientSharingLink = Add-EmailToSharingLink -sharingLink $script:currentClientInfo.SharingLink -email $recipientEmail
    
                $additionalEmailTemplate = Get-EmailTemplate -TemplateName "AdditionalShareLink"
                $additionalEmailBody = Format-EmailTemplate -Template $additionalEmailTemplate -Placeholders $emailPlaceholders
    
                Write-Log "Preparing AdditionalShareLink email for: $recipientEmail"
                $emailsToSend += @{
                    From            = $currentUserEmail
                    To              = @($recipientEmail)
                    Subject         = "Access to $selectedClient Secure Client Portal - Sequoia Tax Relief"
                    Body            = $additionalEmailBody
                    BodyContentType = "Html"
                }
            }
    
            # Prepare "send to self" email if selected
            if ($sendToSelf) {
                $emailPlaceholders = Use-Hash -Hashtable $placeholders
                $emailPlaceholders.ClientSharingLink = $script:currentClientInfo.SharingLink
                $emailPlaceholders.AdditionalEmails = $additionalSelectedEmails -join ', '
                $salesNotificationTemplate = Get-EmailTemplate -TemplateName "SalesResendNotification"
                $salesNotificationBody = Format-EmailTemplate -Template $salesNotificationTemplate -Placeholders $emailPlaceholders
    
                Write-Log "Preparing SalesNotification email to self: $currentUserEmail"
                $emailsToSend += @{
                    From            = $currentUserEmail
                    To              = @($currentUserEmail)
                    Subject         = "A Link Has Been Resent to $selectedClient - Sequoia Tax Relief"
                    Body            = $salesNotificationBody
                    BodyContentType = "Html"
                }
            }
            Show-StatusWindow "Sending emails..."
            # Send all prepared emails
            $allEmailsSent = $true
            foreach ($email in $emailsToSend) {
                try {
                    Send-PnPMailWithTimeout @email
                    Write-Log "Email sent successfully to $($email.To -join ', ')."
                } catch {
                    $allEmailsSent = $false
                    Write-Log "Failed to send email to $($email.To -join ', '): $($_.Exception.Message)" -Level "ERROR"
                }
            }
    
            # Update status based on email sending results
            if ($allEmailsSent) {
                Show-StatusWindow -Close
                # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "All emails sent successfully."
                $sendShareForm.DialogResult = [System.Windows.Forms.MessageBox]::Show("Emails sent successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                Show-StatusWindow -Close
                # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Some emails failed to send."
                $sendShareForm.DialogResult = [System.Windows.Forms.MessageBox]::Show("Some emails failed to send. Please check the logs for details.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
        catch {
            Show-StatusWindow -Close
            $errorMessage = "An error occurred: $($_.Exception.Message)"
            Write-Log $errorMessage -Level "ERROR"
            # Update-StatusLabel -Label $statusLabel -Form $sendShareForm -Status "Error occurred."
            $sendShareForm.DialogResult = [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    
        }
        finally {
            $resendButton.Enabled = $true
        }
    })    
    try {
        Write-Log "Displaying SendShareLink for $env:USERNAME"
        [void]$sendShareForm.ShowDialog()
    } 
    finally {
        Write-Log "$env:USERNAME closed SendShareLink" 
        if ($sendShareForm -and -not $sendShareForm.IsDisposed) {
            Write-Log "Disposing Form"
            $sendShareForm.Dispose()
        }
    }
    if($sendShareForm.DialogResult -eq [System.Windows.Forms.DialogResult]::OK){
        return "Success"
    } else {
        return 2
    }
}
