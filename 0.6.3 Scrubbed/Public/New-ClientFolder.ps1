using namespace System.Windows.Forms
using namespace System.Drawing

# Import required functions
#Import-Module SequoiaTax -Force -Verbose -ErrorAction Stop
function New-ClientFolder{
    [CmdletBinding()]
    param()

    #function that shows the available Team options
    function Show-TeamSelection($title) {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $title
        $form.Size = New-Object System.Drawing.Size(450,120)  # Reduced height since we removed checkboxes
        $form.StartPosition = 'CenterScreen'
        $form.TopMost = $true
        
        $result = $null
    
        # Add Team Buttons
        $buttonX = 20
        $buttonY = 20
        $buttonWidth = 125
        $buttonHeight = 40
        $buttonSpacing = 10 
        foreach ($team in @('TeamAmanda', 'TeamJamie', 'TeamKyle')) {
            $button = New-Object System.Windows.Forms.Button
            $button.Location = New-Object System.Drawing.Point($buttonX,$buttonY)
            $button.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
            $button.Text = $team
            $button.Add_Click({
                $script:result = $this.Text
                $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $form.Close()
            })
            $form.Controls.Add($button)
            $buttonX += $buttonWidth + $buttonSpacing
        }
    
        $form.Add_Shown({$form.Activate()})
        $dialogResult = $form.ShowDialog()
    
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            return @{
                Team = $script:result
                SendToPayments = $true      # Set default values
                SkipTeamEmails = $false      # Set default values
            }
        } else {
            return $null
        }
    }
    
    # function that will reach out to Sharepoint to create the folder and then sends emails
    function Invoke-NewClientFolder {
        param (
        [string]$clientName,
        [string]$clientEmail,
        [string]$folderPath,
        [bool]$isResearchClient,
        [string]$team,
        [bool]$sendToPayments = $false,
        [bool]$skipTeamEmails = $true
    )
        Write-Log "Invoke-NewClientFolder called with:"
        Write-Log "  clientName: $clientName"
        Write-Log "  clientEmail: $clientEmail"
        Write-Log "  folderPath: $folderPath"
        Write-Log "  isResearchClient: $isResearchClient"
        Write-Log "  team: $team"
        Write-Log "  sendToPayments: $sendToPayments"
        Write-Log "  skipTeamEmails: $skipTeamEmails" 

        $scriptBlock = {
            param ($clientName, $clientEmail, $folderPath, $isResearchClient, $team, $sendToPayments, $skipTeamEmails)
            try  {
                
                # Import module with full path
                Import-Module SequoiaTax 
                Import-Module PnP.PowerShell 

                $siteUrlForFolder = Get-SharePointBaseUrl
Connect-PnPOnline

                $currentUserEmail = "$env:USERNAME@sequoiataxrelief.com"

                # Create the folder
                Add-PnPFolder -Name $clientName -Folder $folderPath
                # get folder to share
                $folder = Get-PnPFolder -Url "$folderPath/$clientName"
                #share folder with client
                $sharingResult = Add-PnPFolderUserSharingLink -Folder $folder -ShareType Edit -Users $clientEmail
 
                # Prepare placeholders
                $placeholders = @{
                    ClientName = $clientName
                    CompanyLogoUrl = Get-CompanyUrl -UrlType "Logo"
                    ClientSharingLink = $sharingResult.WebUrl
                    SupportSiteUrl = Get-CompanyUrl -UrlType "Support"
                    PrivacyPolicyUrl = Get-CompanyUrl -UrlType "PrivacyPolicy"
                    CurrentYear = (Get-Date).Year
                    TeamName = $team
                }
    
                if ($isResearchClient) {
                    $clientEmailTemplate = Get-EmailTemplate -TemplateName "ResearchClient"
                    $clientEmailBody = Format-EmailTemplate -Template $clientEmailTemplate -Placeholders $placeholders
                    $clientEmailresult = Send-PnPMailWithTimeout -From $currentUserEmail -To $clientEmail -Subject "Welcome to Your Secure Client Portal - Sequoia Tax Relief" -Body $clientEmailBody
                    Write-Log ($clientEmailResult.Status -ne "Success" ? 
                        "Failed to send client email: $($clientEmailResult.Message | ConvertTo-Json -Compress)" : 
                        "Client email sent successfully, Verification: $($clientEmailResult.VerificationMessage)") -Level ($clientEmailResult.Status -ne "Success" ? "ERROR" : "INFO")

                    $teamEmailTemplate = Get-EmailTemplate -TemplateName "ResearchClientTeam"
                    $teamEmailBody = Format-EmailTemplate -Template $teamEmailTemplate -Placeholders $placeholders
                    $teamUsers = Get-PnPGroupMember -Identity $team | Select-Object -ExpandProperty Email
                    $cleanTeamUsers = @($currentUserEmail) + $teamUsers | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
                    $teamEmailResult = Send-PnPMailWithTimeout -From $currentUserEmail -To $cleanTeamUsers -Subject "New Research Client Folder Created - Action Required" -Body $teamEmailBody
                    Write-Log ($teamEmailResult.Status -ne "Success" ? 
                        "Failed to send team email: $($teamEmailResult.Message | ConvertTo-Json -Compress)" : 
                        "Team email sent successfully, Verification: $($teamEmailResult.VerificationMessage)") -Level ($teamEmailResult.Status -ne "Success" ? "ERROR" : "INFO")
                    if ($sendToPayments) {
                        $paymentsEmailTemplate = Get-EmailTemplate -TemplateName "ResearchClientPayments"
                        $paymentsEmailBody = Format-EmailTemplate -Template $paymentsEmailTemplate -Placeholders $placeholders
                        $paymentsEmailResult = Send-PnPMailWithTimeout -From $currentUserEmail -To "payments@sequoiataxrelief.com" -Subject "New Research Client Folder Created - Payment Processing Required" -Body $paymentsEmailBody -BodyContentType Html
                        Write-Log ($paymentsEmailResult.Status -ne "Success" ? 
                            "Failed to send payments email: $($paymentsEmailResult.Message | ConvertTo-Json -Compress)" : 
                            "Payments email sent successfully, Verification: $($paymentsEmailResult.VerificationMessage)") -Level ($paymentsEmailResult.Status -ne "Success" ? "ERROR" : "INFO")
                    }
                } else {
                    # send general client email
                    $clientEmailTemplate = Get-EmailTemplate -TemplateName "GeneralClient"
                    $clientEmailBody = Format-EmailTemplate -Template $clientEmailTemplate -Placeholders $placeholders
                    $clientEmailResult = Send-PnPMailWithTimeout -From $currentUserEmail -To $clientEmail -Subject "Welcome to Your Secure Client Portal - Sequoia Tax Relief" -Body $clientEmailBody -BodyContentType Html
                    Write-Log ($clientEmailResult.Status -ne "Success" ? 
                        "Failed to send client email: $($clientEmailResult.Message | ConvertTo-Json -Compress)" : 
                        "Client email sent successfully, Verification: $($clientEmailResult.VerificationMessage)") -Level ($clientEmailResult.Status -ne "Success" ? "ERROR" : "INFO")
                    # send general team email
                    $generalTeamEmailTemplate = Get-EmailTemplate -TemplateName "GeneralClientTeam"
                    $generalTeamEmailBody = Format-EmailTemplate -Template $generalTeamEmailTemplate -Placeholders $placeholders
                    $teamEmailResult = Send-PnPMailWithTimeout -From $currentUserEmail -To $currentUserEmail -Subject "New General Client Folder Created" -Body $generalTeamEmailBody -BodyContentType Html
                    Write-Log ($teamEmailResult.Status -ne "Success" ? 
                        "Failed to send team email: $($teamEmailResult.Message | ConvertTo-Json -Compress)" : 
                        "Team email sent successfully, Verification: $($teamEmailResult.VerificationMessage)") -Level ($teamEmailResult.Status -ne "Success" ? "ERROR" : "INFO")
                }
                return @{
                    Status = "Success"
                    SharingLink = $sharingResult.WebUrl
                    FolderUrl = "$folderPath/$clientName"
                }
            }
            catch {
                $errorResult = @{
                    Status = "Error"
                    ErrorMessage = $_.Exception.Message
                }
                Write-Error "Error in Invoke-NewClientFolder: $_"
                Write-Verbose "Operation failed. Result: $($errorResult | ConvertTo-Json -Compress)"
                return $errorResult
            }
            finally {
                Write-Verbose "Removing PnP connection"
                Remove-PnPConnection
            }
        }

        $powershell = [powershell]::Create().AddScript($scriptBlock).AddParameters(@{
            clientName = $clientName
            clientEmail = $clientEmail
            folderPath = $folderPath
            isResearchClient = $isResearchClient
            team = $team
            sendToPayments = $sendToPayments
            skipTeamEmails = $skipTeamEmails
            LOGFILE = $LOGFILE
        })

        $powershell.RunspacePool = $script:appInstance.RunspacePool
        Write-Log "Starting asynchronous job"
        $asyncResult = $powershell.BeginInvoke()

        Add-CleanupItem -Item $powershell -Type "Job"
        return @{
            PowerShell = $powershell
            AsyncResult = $asyncResult
        }
    }
    # create the mainForm for the New-ClientFolder Menu
    # Main form
    $mainForm = New-StandardForm 'Create Client Share' 300 280

    # Client Name
    $mainForm.Controls.Add((New-Label 'Client Name:' 10 20 100 20))
    $clientNameTextBox = New-TextBox 120 20 150 20
    $clientNameTextBox.Add_Leave({
        $clientName = $this.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($clientName)) {
            if (Test-ValidEmailFormat -Email $clientName) {
            Show-TopMostMessageBox "Please do not enter an email address as the clients name." "Invalid Name" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Warning)
            $this.Text = ""
            $this.Focus()
            return
            }
        }
    })
    $mainForm.Controls.Add($clientNameTextBox)
    
    # Client Email
    $mainForm.Controls.Add((New-Label 'Client Email:' 10 50 100 20))
    $clientEmailTextBox = New-TextBox 120 50 150 20
    $clientEmailTextBox.Add_Leave({
        $email = $this.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($email)){
            if (-not (Test-ValidEmailFormat -Email $email)) {
                Show-TopMostMessageBox "Please enter a valid email address." "Invalid Email" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Warning)
                $this.Focus()
                return
            }
        }
    })
    $mainForm.Controls.Add($clientEmailTextBox)
    
    # Status Label
    $statusLabel = New-Label "Status: Ready" 10 140 280 20
    $mainForm.Controls.Add($statusLabel)
    
    $result = "Cancelled" # Default result in case of errors and shit
    # General Button
    $generalButton = New-Button 'General' 50 160 200 30 {
        $result = ProcessClientFolder -IsResearch $false
        if ($result -ne "Cancelled") {
            $mainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $mainForm.Close()
        }
    }
    $mainForm.Controls.Add($generalButton)

    # Research Client Button
    $researchButton = New-Button 'Research Client' 50 190 200 30 {
        $result = ProcessClientFolder -IsResearch $true
        if ($result -ne "Cancelled") {
            $mainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $mainForm.Close()
        }
    }
    $mainForm.Controls.Add($researchButton)
    function ProcessClientFolder {
        param ([bool]$IsResearch)

        $clientName = $clientNameTextBox.Text.Trim()
        $clientEmail = $clientEmailTextBox.Text.Trim()

        # make sure the client email name are not blank
        if ([string]::IsNullOrWhiteSpace($clientName) -or [string]::IsNullOrWhiteSpace($clientEmail)) {
            Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Client Name & Email Are Required"
            Show-TopMostMessageBox "Please enter both client name and email." "Missing Information" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Warning)
            return "Cancelled"
        }
        # Double check email format
        if (-not (Test-ValidEmailFormat -Email $clientEmail)) {
            Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Invalid Email Format"
            Show-TopMostMessageBox "Please enter a valid email address." "Invalid Email" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Warning)
            $clientEmailTextBox.Focus()
            return "Cancelled"
        }
        # make sure the client name is not an email address for Fred
        # Double check client name
        if (Test-ValidEmailFormat -Email $clientName) {
            Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Invalid Client Name"
            Show-TopMostMessageBox "Client name cannot be an email address." "Invalid Input" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Warning)
            $clientNameTextBox.Focus()
            return "Cancelled"
        }

        $confirmResult = Show-TrickyMessageBox -ClientName $clientName -ClientEmail $clientEmail -Title "Confirm $($IsResearch ? 'Research' : 'General') Client Information"
        if (-not $confirmResult) {
            Write-Log "Client Information validation failed or was canceled"
            Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Operation cancelled. Please try again."
            return "Cancelled"
        }

        $team = $null
        $sendToPayments = $true
        $skipTeamEmails = $false
        $folderPath = "Shared Documents/General/TeamSales"

        if ($IsResearch) {
            $teamSelection = Show-TeamSelection 'Select Team for Research Clients'
            if ($null -eq $teamSelection) {
                Write-Log "Operation cancelled. No team selected." -Level "WARN"
                Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Operation cancelled. No team selected."
                Show-TopMostMessageBox "No team selected." "Cancelled" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Information)
                return "Cancelled"
            }
            $team = $teamSelection.Team
            $sendToPayments = $teamSelection.SendToPayments
            $skipTeamEmails = $teamSelection.SkipTeamEmails
            $folderPath = "Shared Documents/Research Clients/$team"
        }

        Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Creating $($IsResearch ? 'Research' : 'General') folder..."
        Show-StatusWindow "Creating $($IsResearch ? 'Research' : 'General') Folder..."
        $job = Invoke-NewClientFolder $clientName $clientEmail $folderPath $IsResearch $team $sendToPayments $skipTeamEmails $localRoot $logFile
        try {
            Write-Log "Waiting for job to complete..."
            $completed = $job.AsyncResult.AsyncWaitHandle.WaitOne(300000) # 5-minute timeout
            
            if ($completed) {
                $jobResult = $job.PowerShell.EndInvoke($job.AsyncResult)
                Write-Log "Job result retrieved: $jobResult"
                
                if ($null -eq $jobResult) {
                    throw "Job result is null"
                }

                $resultHashtable = $jobResult[1]  # We're only using the hashtable part of the result
                Write-Log "Result Hashtable: $($resultHashtable | ConvertTo-Json -Compress)"
                Show-StatusWindow "Adding client to cache..."
                if ($resultHashtable.Status -eq "Success") {
                    Write-Log "Operation successful. Updating UI..."
                    Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Folder created and shared successfully."
                     # Update main form tag
                    $mainForm.Tag = @{
                        Result = "Success"
                        ClientName = $clientName
                        ClientEmail = $clientEmail
                        FolderType = if ($IsResearch) { "Research" } else { "General" }
                        Team = $team
                        SharingLink = $resultHashtable.SharingLink
                    }
                    Write-Log "Main form tag set: $($mainForm.Tag | ConvertTo-Json -Compress)"

                    # Update JSON cache
                    try {
                        Write-Log "Adding client to JSON cache..."
                        $params = @{
                            clientName = $clientName
                            clientEmails = @($clientEmail)
                            sharingLink = $resultHashtable.SharingLink
                            url = $resultHashtable.FolderUrl
                            type = "Folder"
                            isResearchClient = $IsResearch
                            teamName = $team
                        }
                        Write-Log "Calling New-ClientFolderJsonEntry with parameters: $($params | ConvertTo-Json -Compress)"
                        New-ClientFolderJsonEntry @params
                        Write-Log "Successfully added client to JSON cache"
                    } catch {
                        Write-Log "Error adding client to JSON cache: $($_.Exception.Message)" -Level "ERROR"
                        Write-Log "Error details: $($_.Exception.GetType().FullName)" -Level "ERROR"
                        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
                        # Don't throw here, continue with the process
                    }
                    # Show success message to user
                    Show-TopMostMessageBox "Folder created and shared successfully. Emails sent." "Success" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Information)
                    Show-StatusWindow -Close
                    return "Success"
                } else {
                    throw "Operation failed: $($resultHashtable.Status)"
                }
            } else {
                throw "Operation timed out after 5 minutes"
            }
        }
        catch {
            Write-Log "Error in ProcessClientFolder: $_" -Level "ERROR"
            Write-Log "Error details: $($_.Exception.GetType().FullName)" -Level "ERROR"
            Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
            Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Operation failed. Please check the logs."
            Show-StatusWindow "Error in ProcessClientFolder"
            Show-TopMostMessageBox "Operation failed. Please check the logs." "Error" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error)
            Start-Sleep -Seconds 5
            Show-StatusWindow -Close
            return "Error"
        }
        finally {
            Write-Log "Entering finally block..."
            if ($null -ne $job -and $null -ne $job.PowerShell) {
                if (-not [string]::IsNullOrWhiteSpace($job.PowerShell.Streams.Information)) {
                    Write-Log "Job information output: $($job.PowerShell.Streams.Information | ConvertTo-Json -Compress -Depth 10)"
                }
                if (-not [string]::IsNullOrWhiteSpace($job.PowerShell.Streams.Error)) {
                    Write-Log "Job error output: $($job.PowerShell.Streams.Error | ConvertTo-Json -Compress -Depth 10)"
                }
                Remove-CleanupItem -Item $job.PowerShell -Type "Job"
                $job.PowerShell.Dispose()
            } else {
                Write-Log "Job or PowerShell object is null"
            }
            if ($script:statusForm -and $script:statusForm.Visible) {
                Show-StatusWindow -Close
            }
            Write-Log "Exiting ProcessClientFolder"
        }
    }
    try {
        Write-Log "Displaying CreateFolder for $env:USERNAME"
        $dialogResult = $mainForm.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $operationResult = $mainForm.Tag
            Write-Log "CreateFolder completed successfully: $($operationResult | ConvertTo-Json -Compress)"
            return $operationResult
        } else {
            Write-Log "CreateFolder was cancelled or closed"
            return "Cancelled"
        }
    }
    finally {
        Write-Log "$env:USERNAME closed CreateFolder"
        Invoke-Cleanup -Form $mainForm
        $mainForm.Close()
        $mainForm.Dispose()
    }
}
