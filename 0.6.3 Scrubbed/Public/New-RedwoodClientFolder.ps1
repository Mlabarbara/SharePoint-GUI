using namespace System.Windows.Forms
using namespace System.Drawing

# Import required functions
. (Join-Path $PSScriptRoot '..\Public\UI\New-StandardForm.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-Label.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-TextBox.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-Button.ps1')
. (Join-Path $PSScriptRoot '..\Public\UI\New-ComboBox.ps1')
. (Join-Path $PSScriptRoot '..\Private\Utility\Write-Log.ps1')

function New-RedwoodClientFolder {
    [CmdletBinding()]
    param()

    function Invoke-NewClientFolder($clientName, $clientEmail) {
        Write-Log "Invoke-NewClientFolder called with:"
        Write-Log "  clientName: $clientName"
        Write-Log "  clientEmail: $clientEmail"
        
        $scriptBlock = {
            param ($clientName, $clientEmail)
            try {
                Import-Module "\\str-0111\MainMenu\0.6.2-MoreCache\SharePointModule.psm1" -Force -Verbose
                Import-Module PnP.PowerShell
                $siteUrl = "https://sequoiataxrelief.sharepoint.com/sites/redwoodtaxsvcs"
Connect-PnPOnline

                $currentUserEmail = "geoff@redwoodtaxsvcs.com"
                $folderPath = "Shared Documents/Client Folders"
                
                # Create the folder
                Add-PnPFolder -Name $clientName -Folder $folderPath
                # get newly created folder
                $folder = Get-PnPFolder -Url $folderPath/$clientName
                
                #share folder with client
                $sharingResult = Add-PnPFolderUserSharingLink -Folder $folder -Users $clientEmail
 
                # Prepare placeholders
                $placeholders = @{
                    ClientName = $clientName
                    CompanyLogoUrl = Get-CompanyUrl -UrlType "Logo"
                    ClientSharingLink = $sharingResult.WebUrl
                    SupportSiteUrl = Get-CompanyUrl -UrlType "Support"
                    PrivacyPolicyUrl = Get-CompanyUrl -UrlType "PrivacyPolicy"
                    CurrentYear = (Get-Date).Year
                }
    
                # Send client email
                $clientEmailTemplate = Get-EmailTemplate -TemplateName "RedwoodClient"
                $clientEmailBody = Format-EmailTemplate -Template $clientEmailTemplate -Placeholders $placeholders
                Send-PnPMail -From $currentUserEmail -To $clientEmail -Subject "Welcome to Your Secure Client Portal - Redwood Tax Services" -Body $clientEmailBody -BodyContentType Html
    
                $generalTeamEmailTemplate = Get-EmailTemplate -TemplateName "GeneralClientTeam"
                $generalTeamEmailBody = Format-EmailTemplate -Template $generalTeamEmailTemplate -Placeholders $placeholders
                Send-PnPMail -From $currentUserEmail -To $currentUserEmail -Subject "New Client Folder Created" -Body $generalTeamEmailBody -BodyContentType Html
            
                return @{
                    Status = "Success"
                    SharingLink = $sharingResult.WebUrl
                    FolderUrl = $folder.ServerRelativeUrl
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

    # Main form
    $mainForm = New-StandardForm 'Create Client Folder' 300 220

    # Client Name
    $mainForm.Controls.Add((New-Label 'Client Name:' 10 20 100 20))
    $clientNameTextBox = New-TextBox 120 20 150 20
    $mainForm.Controls.Add($clientNameTextBox)
    
    # Client Email
    $mainForm.Controls.Add((New-Label 'Client Email:' 10 50 100 20))
    $clientEmailTextBox = New-TextBox 120 50 150 20
    $mainForm.Controls.Add($clientEmailTextBox)
    
    # Status Label
    $statusLabel = New-Label "Status: Ready" 10 110 280 20
    $mainForm.Controls.Add($statusLabel)
    
    $result = "Cancelled" # Default result in case of errors
    # Create Folder Button
    $createButton = New-Button 'Create Folder' 50 140 200 30 {
        $result = ProcessClientFolder
        if ($result -ne "Cancelled") {
            $mainForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $mainForm.Close()
        }
    }
    $mainForm.Controls.Add($createButton)

    # Add form closing event
    $mainForm.Add_FormClosing({
        param($sender1, $e)
        Write-Log "Form closing event triggered"
        if ($e.CloseReason -ne [System.Windows.Forms.CloseReason]::None) {
            Write-Log "Form is closing. Performing cleanup..."
        }
    })
    function New-RedwoodCacheEntry{
        param(
            [string]$clientName,
            [string[]]$clientEmail,
            [string]$sharingLink,
            [string]$url
        )
        #get the cache file path
        $jsonPath = Get-CacheFilePath -CacheType "Redwood"
        Write-Log "Adding new client to redwood cache: $clientName | $clientEmail | $sharingLink | $url"
        try {
            #check if the cache file exists
            if (Test-Path $jsonPath){
                $cachedData = Get-Content $jsonPath | ConvertFrom-Json
                if (-not ($cachedData -is [array])) {
                    $cachedData = @($cachedData)
                }
            } else {
                $cachedData = @{}
            }
            # create a new client object
            $newClient = New-FolderClientObject -clientName $clientName `
                                                -clientEmails $clientEmail `
                                                -sharingLink $sharingLink `
                                                -url $url `
                                                -location "Redwood"                                
            # add the new client to the cache
            $cachedData += $newClient
            # sort clients by name
            $cachedData = $cachedData | Sort-Object -Property ClientName
            # write updated cache to file
            $cachedData | ConvertTo-Json -Depth 10 | Set-Content $jsonPath
            Write-Log "Successfully added client to redwood cache: $clientName | $clientEmail | $sharingLink | $url"
        } catch {
            Write-Log "Error adding client to redwood cache: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    function ProcessClientFolder {
        $clientName = $clientNameTextBox.Text
        $clientEmail = $clientEmailTextBox.Text

        if ([string]::IsNullOrWhiteSpace($clientName) -or [string]::IsNullOrWhiteSpace($clientEmail)) {
            Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Client Name & Email Are Required"
            Show-TopMostMessageBox "Please enter both client name and email." "Missing Information" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Warning)
            return "Cancelled"
        }

        $confirmResult = Show-TrickyMessageBox -ClientName $clientName -ClientEmail $clientEmail -Title "Confirm Client Information"
        if (-not $confirmResult) {
            Write-Log "Client Information validation failed or was canceled"
            Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Operation cancelled. Please try again."
            return "Cancelled"
        }

        Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Creating folder..."
        Show-StatusWindow "Creating Client Folder..."
        $job = Invoke-NewClientFolder $clientName $clientEmail

        try {
            Write-Log "Waiting for job to complete..."
            $completed = $job.AsyncResult.AsyncWaitHandle.WaitOne(300000) # 5-minute timeout
            
            if ($completed) {
                $jobResult = $job.PowerShell.EndInvoke($job.AsyncResult)
                Write-Log "Job result retrieved: $jobResult"
                
                if ($null -eq $jobResult) {
                    throw "Job result is null"
                }

                $resultHashtable = $jobResult[1]
                Write-Log "Result Hashtable: $($resultHashtable | ConvertTo-Json -Compress)"
    
                if ($resultHashtable.Status -eq "Success") {
                    Write-Log "Operation successful. Updating UI..."
                    Update-StatusLabel -Label $statusLabel -Form $mainForm -Status "Folder created and shared successfully."
                    Show-StatusWindow -Close
                    # Update main form tag
                    $mainForm.Tag = @{
                        Result = "Success"
                        ClientName = $clientName
                        ClientEmail = $clientEmail
                        SharingLink = $resultHashtable.SharingLink
                        Url = $resultHashtable.FolderUrl
                    }
                    Write-Log "Main form tag set: $($mainForm.Tag | ConvertTo-Json -Compress)"
                    
                    #update the json Cache
                    try {
                        New-RedwoodCacheEntry -clientName $clientName -clientEmail $clientEmail -sharingLink $resultHashtable.SharingLink -url $resultHashtable.FolderUrl
                    } catch {
                        Write-Log "Error adding client to redwood cache: $($_.Exception.Message)" -Level "ERROR"
                        Write-Log "Error details: $($_.Exception.GetType().FullName)" -Level "ERROR"
                        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
                        # Don't throw here, continue with the process
                    }
                    
                    # Show success message to user
                    Show-TopMostMessageBox "Folder created and shared successfully. Emails sent." "Success" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Information)
                    
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
                    Write-Log "Job information output: $($job.PowerShell.Streams.Information | Out-String)"
                }
                if (-not [string]::IsNullOrWhiteSpace($job.PowerShell.Streams.Error)) {
                    Write-Log "Job error output: $($job.PowerShell.Streams.Error | Out-String)"
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
