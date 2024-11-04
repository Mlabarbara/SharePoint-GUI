using namespace System.Windows.Forms
using namespace System.Drawing

# Import required functions
# . (Join-Path $PSScriptRoot '..\Public\UI\New-StandardForm.ps1')
# . (Join-Path $PSScriptRoot '..\Public\UI\New-Label.ps1')
# . (Join-Path $PSScriptRoot '..\Public\UI\New-TextBox.ps1')
# . (Join-Path $PSScriptRoot '..\Public\UI\New-Button.ps1')
# . (Join-Path $PSScriptRoot '..\Public\UI\New-ComboBox.ps1')
# . (Join-Path $PSScriptRoot '..\Private\Utility\Write-Log.ps1')

#end New-ClientFolder   
function Show-ConvertGeneralToResearchForm {
    [CmdletBinding()]
    param()
    # Connect to SharePoint and get client list
    Show-StatusWindow "Getting TeamSales Clients"
    $cachedData = Get-ClientJsonCache -selectedTeam "TeamSales"
    $clients = $cachedData.General | Select-Object -ExpandProperty ClientName | Sort-Object
    Show-StatusWindow -Close

    # Create and show the main form
    $convertForm = New-StandardForm 'Convert General Folder to Research Client' 520 300

    # Client dropdown
    $convertForm.Controls.Add((New-Label 'Select Client:' 10 20 80 20))
    $clientComboBox = New-ComboBox 120 20 330 20 $clients
    $convertForm.Controls.Add($clientComboBox)
    $convertForm.Controls.Add((New-Label 'Select Team:' 10 50 80 20))

    $teamAmandaCheckbox = New-Object System.Windows.Forms.CheckBox
    $teamAmandaCheckbox.Text = "TeamAmanda"
    $teamAmandaCheckbox.Location = New-Object System.Drawing.Point(120, 50)
    $teamAmandaCheckbox.Size = New-Object System.Drawing.Size(120, 20)
    $convertForm.Controls.Add($teamAmandaCheckbox)

    $teamJamieCheckbox = New-Object System.Windows.Forms.CheckBox
    $teamJamieCheckbox.Text = "TeamJamie"
    $teamJamieCheckbox.Location = New-Object System.Drawing.Point(250, 50)
    $teamJamieCheckbox.Size = New-Object System.Drawing.Size(120, 20)
    $convertForm.Controls.Add($teamJamieCheckbox)

    $teamKyleCheckbox = New-Object System.Windows.Forms.CheckBox
    $teamKyleCheckbox.Text = "TeamKyle"
    $teamKyleCheckbox.Location = New-Object System.Drawing.Point(370, 50)
    $teamKyleCheckbox.Size = New-Object System.Drawing.Size(120, 20)
    $convertForm.Controls.Add($teamKyleCheckbox)

    # Checkboxes
    $sendPaymentsCheckbox = New-Object System.Windows.Forms.CheckBox
    $sendPaymentsCheckbox.Text = "Send Email To Payments?"
    $sendPaymentsCheckbox.Location = New-Object System.Drawing.Point(10, 80)
    $sendPaymentsCheckbox.Size = New-Object System.Drawing.Size(200, 20)

    if ($env:USERNAME -eq "marklabarbara") {
        $sendPaymentsCheckbox.Checked = $false
        $sendPaymentsCheckbox.Enabled = $true
    } else {
        $sendPaymentsCheckbox.Checked = $true
        $sendPaymentsCheckbox.Enabled = $false
    }

    $convertForm.Controls.Add($sendPaymentsCheckbox)

    $sendTeamCheckbox = New-Object System.Windows.Forms.CheckBox
    $sendTeamCheckbox.Text = "Send Email To Team?"
    $sendTeamCheckbox.Location = New-Object System.Drawing.Point(10, 110)
    $sendTeamCheckbox.Size = New-Object System.Drawing.Size(200, 20)
    $convertForm.Controls.Add($sendTeamCheckbox)

    # Status label
    $statusLabel = New-Label "Status: Ready" 10 140 380 20
    $convertForm.Controls.Add($statusLabel)
    # Convert button
    
    $convertButton = New-Button 'Convert General to Research Client' 175 170 220 50 {
        $selectedClient = $clientComboBox.SelectedItem
        $selectedTeam = $null
        $selectedTeamCount = @($teamAmandaCheckbox.Checked, $teamJamieCheckbox.Checked, $teamKyleCheckbox.Checked).Where({$_ -eq $true}).Count

        if ($selectedTeamCount -gt 1) {
            [System.Windows.Forms.MessageBox]::Show("Please select only one team.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        } elseif ($selectedTeamCount -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a team.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        } else {
            if ($teamAmandaCheckbox.Checked) { $selectedTeam = "TeamAmanda" }
            elseif ($teamJamieCheckbox.Checked) { $selectedTeam = "TeamJamie" }
            elseif ($teamKyleCheckbox.Checked) { $selectedTeam = "TeamKyle" }
        }

        $sendPayments = $sendPaymentsCheckbox.Checked
        $sendTeam = $sendTeamCheckbox.Checked
    
        if ([string]::IsNullOrWhiteSpace($selectedClient) -or [string]::IsNullOrWhiteSpace($selectedTeam)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a client and a team.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    
        # Prepare confirmation message
        $confirmMessage = "Please confirm the following details:`n`n"
        $confirmMessage += "Client: $selectedClient`n"
        $confirmMessage += "Team: $selectedTeam`n`n"
        $confirmMessage += "Emails will be sent to:`n"
        if ($sendPayments) { $confirmMessage += "- Payments Department`n" }
        if ($sendTeam) { $confirmMessage += "- Selected team members`n" }
        
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            $confirmMessage,
            "Confirm Conversion",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
    
        if ($confirmResult -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }
    
        try {
            Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Connecting to Sharepoint Online"

Connect-PnPOnline
            $siteUrlBase = Get-SharePointBaseUrl
            $currentUserEmail = "$env:USERNAME@sequoiataxrelief.com"
            $emailsSent = $true
            
            # Move folder
            Write-Log "Moving client folder from $srouceSiteUrl, client: $selectedClient, team: $selectedTeam"
            $moveResult = Move-ClientFolder -sourceSiteUrl $siteUrlBase -clientName $selectedClient -teamName $selectedTeam
    
            if (-not $moveResult) {
                throw "Failed to move client folder for $selectedTeam"
            }
    
            # Get sharing info
            Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Getting Folder Sharing Link for $selectedTeam"
            $sharingInfo = Get-FolderSharingInfo -siteUrl $siteUrlBase -folderPath "Shared Documents/Research Clients/$selectedTeam/$selectedClient"
            Write-Log "Sharing Info: $($sharingInfo | ConvertTo-Json -Depth 1 -Compress)"
            if ($null -eq $sharingInfo) {
                throw "Failed to get folder sharing information for $selectedTeam"
            }
            Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Folder moved successfully for $selectedTeam"
            
            $clientEmails = $sharingInfo.GrantedIdentities
            if ($null -eq $clientEmails -or $clientEmails.Count -eq 0) {
                Write-Log "Warning: No client emails found in sharing info for selected client: $selectedClient" -Level "WARN"
                $clientEmails = @()
            }
            $newUrl = "/sites/secureclientupload/Shared Documents/Research Clients/$selectedTeam/$selectedClient"
            # update the client cache json file:
            try {
                Move-ClientFolderJson -clientName $selectedClient -newTeam $selectedTeam -newSharingLink $sharingInfo.SharingLink -newUrl $newUrl -clientEmails $clientEmails
                Write-Log "Updated folder cache for client: $selectedClient"
#                Update-ClientJsonCache -teamName $selectedTeam -clientName $selectedClient -email $clientEmails -sharingLink $sharingInfo.SharingLink -isNewSubsite:$false
#                Write-Log "Updated JSON cache for client: $selectedClient with emails: $($clientEmails -join ', ')"
                Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Client cache updated successfully"
            } catch {
                Write-Log "Error updating folder cache: $($_.Exception.Message)" -Level "ERROR"
                Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Error updating client cache"
                # Note: We're not throwing an exception here to allow the process to continue even if cache update fails
            }

            # Prepare email placeholders
            $placeholders = @{
                ClientName = $selectedClient
                CompanyLogoUrl = Get-CompanyUrl -UrlType "Logo"
                ClientSharingLink = if ($sharingInfo.SharingLink -is [array]){
                    Write-Log "Sharing Link is an array, using first link" -Level "WARN"
                    $sharingInfo.SharingLink[0]
                } else {
                    $sharingInfo.SharingLink
                }
                SupportSiteUrl = Get-CompanyUrl -UrlType "Support"
                PrivacyPolicyUrl = Get-CompanyUrl -UrlType "PrivacyPolicy"
                CurrentYear = (Get-Date).Year
                TeamName = $selectedTeam
            }
    
            # Send team email
            if ($sendTeam) {
                try {
                    Write-Log "Preparing team email for $selectedTeam"
                    Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Preparing team email for $selectedTeam"
                    Write-Log "Placeholders for team email: $($placeholders | ConvertTo-Json -Depth 1)"
                    $teamEmailTemplate = Get-EmailTemplate -TemplateName "ResearchClientTeam"
                    $teamEmailBody = Format-EmailTemplate -Template $teamEmailTemplate -Placeholders $placeholders
                    $teamUsers = Get-PnPGroupMember -Identity $selectedTeam -ErrorAction Stop | Select-Object -ExpandProperty Email
                    if ($null -eq $teamUsers -or $teamUsers.Count -eq 0) {
                        throw "No team members found for $selectedTeam"
                    }
                    $teamEmailRecipients = @($currentUserEmail) + $teamUsers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
                    Write-Log "Sending team email to: $($teamEmailRecipients -join ', ')"
                    Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Sending team email for $selectedTeam"
                    Send-PnPMailWithTimeout -From $currentUserEmail -To $teamEmailRecipients -Subject "New Research Client Folder Created - Action Required" -Body $teamEmailBody -BodyContentType Html
                    Write-Log "Team email sent successfully for $selectedTeam"
                    Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Team email sent successfully for $selectedTeam"
                }
                catch {
                    $emailsSent = $false
                    Write-Log "Error sending team email for $($selectedTeam): $_" -Level "ERROR"
                    Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Error sending team email for $selectedTeam"
                }
            }
    
            # Send payments email
            if ($sendPayments) {
                try {
                    Write-Log "Preparing payments email"
                    $paymentsEmailTemplate = Get-EmailTemplate -TemplateName "ResearchClientPayments"
                    $paymentsEmailBody = Format-EmailTemplate -Template $paymentsEmailTemplate -Placeholders $placeholders
                    Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Sending Payments Email"
                    Send-PnPMailWithTimeout -From $currentUserEmail -To "payments@sequoiataxrelief.com" -Subject "New Research Client Folder Created - Payment Processing Required" -Body $paymentsEmailBody -BodyContentType Html
                    Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Payments Email Sent"
                }
                catch {
                    $emailsSent = $false
                    Write-Log "Error sending payments email: $_" -Level "ERROR"
                    Update-StatusLabel -Label $statusLabel -Form $convertForm -Status "Error sending payments email"
                    Start-Sleep -Seconds 5
                    $convertForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                }
            }
        
            if ($emailsSent) {
                $finalMessage = "All tasks completed successfully:`n"
                $finalMessage += "- Client folder moved to: $selectedTeam`n"
                if ($sendPayments) { $finalMessage += "- Payments email sent`n" }
                if ($sendTeam) { $finalMessage += "- Team email sent to: $selectedTeam`n" }
                $result = [System.Windows.Forms.MessageBox]::Show(
                    $finalMessage,
                    "Operation Completed",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                Write-Log "All emails sent successfully, result: $result"
                $convertForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $convertForm.Close()
            }
            else {
                Write-Log "One or more emails failed to send. Please check the logs for details." -Level "ERROR"
                Start-Sleep -Seconds 3
                $convertForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $convertForm.Close()
            }
        }
        catch {
            Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "An error occurred: $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            $convertForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        } 
    }
    $convertForm.Controls.Add($convertButton)

    # Form closing event
    $cleanupPerformed = $false
    $convertForm.Add_FormClosing({
        param($sendClose, $e)
        if (-not $cleanupPerformed) {
            Write-Log "ConvertFolderForm is closing. Performing cleanup..."
            $convertForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            Invoke-Cleanup -Form $convertForm
            $cleanupPerformed = $true
        }
    })
    try {# Show the form
        Write-Log "Displaying ConvertFolderForm for $env:USERNAME"
        [void]$convertForm.ShowDialog()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred while loading: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Log "Error in Show-ConvertGeneralToResearch: $_" -Level "ERROR"
    } finally {
        Write-Log "$env:USERNAME closed ConvertFolderForm"
    }
}
