using namespace System.Windows.Forms
using namespace System.Drawing

function New-ClientSubsite { 
    <#
    .SYNOPSIS
    Creates a new client subsite and sends emails.

    .DESCRIPTION
    This function creates a new client subsite in SharePoint, sets up necessary configurations, and sends notification emails. It optionally allows specifying the sender's email address.

    .PARAMETER clientName
    The name of the client.

    .PARAMETER clientEmail
    The email address of the client.

    .PARAMETER TeamName
    The team associated with the client.

    .PARAMETER EmailTemplateName
    The name of the email template to use.

    .PARAMETER skipAlerts
    If specified, alerts creation is skipped.

    .PARAMETER sendPaymentsEmail
    If specified, an email is sent to payments.

    .PARAMETER From
    The email address to use as the sender. If not provided, defaults to the current user's email.

    .EXAMPLE
    New-ClientSubsite -clientName "ABC Corp" -clientEmail "contact@abccorp.com" -TeamName "TeamJamie" -From "Jeff@sequoiataxrelief.com"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$clientName,
        [Parameter(Mandatory=$true)]
        [string]$clientEmail,
        [Parameter(Mandatory=$true)]
        [string]$TeamName,
        [Parameter()]
        [string]$EmailTemplateName = "ClientWelcome",
        [switch]$skipAlerts,
        [switch]$skipEmails,
        [switch]$sendPaymentsEmail,  
        [string]$From
    )
    begin {
        Import-Module PnP.PowerShell
        Import-Module "$script:LOCAL_ROOT\HelperFunctions.psm1" -Force
        Import-Module "$script:LOCAL_ROOT\SharePointModule.psm1" -Force
    }
    process {
        Write-Log "Starting New-ClientSubsite function"
        Write-Log "Parameters: clientName=$clientName, clientEmail=$clientEmail, TeamName=$TeamName, skipAlerts=$skipAlerts, skipEmails=$skipEmails, sendPaymentsEmail=$sendPaymentsEmail, From=$From"

        $siteUrlBase = Get-SharePointBaseUrl
        try {
            $siteUrl = "$siteUrlBase/$TeamName"
            Write-Log "Connecting to SharePoint site: $siteUrl"
            Connect-PnPOnline -Url $siteUrl -ClientId 66169a0f-1b2f-49af-acba-024033664ec1 `
                -Tenant "sequoiataxrelief.com" `
Connect-PnPOnline
          
            if (Get-PnPContext) {
                Write-Log "Successfully connected to SharePoint"
            } else {
                throw "Failed to establish SharePoint connection"
            }

            Write-Log "Getting PnP Web information"
            $currentWeb = Get-PnPWeb
            $currentWebServerRelativeUrl = $currentWeb.ServerRelativeUrl
            Write-Log "Current web server relative URL: $currentWebServerRelativeUrl"

            $subsiteUrl = $clientName -replace '\s', '' -replace '[^a-zA-Z0-9]', ''
            $fullSubsiteUrl = "$currentWebServerRelativeUrl/$subsiteUrl"
            Write-Log "Full subsite URL: $fullSubsiteUrl"
            $subsite = Get-PnPSubWeb -Identity $fullSubsiteUrl -ErrorAction SilentlyContinue

            if ($null -eq $subsite) {
                Write-Log "Subsite does not exist. Creating subsite..."
                Show-StatusWindow "Creating client subsite..."
                Write-Log "Trying to create Title: $clientName subsiteUrl: $subsiteUrl"
                New-PnPWeb -Title $clientName -Url $subsiteUrl -Template "BDR#0" -Locale 1033 -ErrorAction Stop
                Write-Log "Trying to connect to $siteUrl/$subsiteUrl"
                Connect-PnPOnline -Url "$siteUrl/$subsiteUrl" -ClientId 66169a0f-1b2f-49af-acba-024033664ec1 `
                    -Tenant "sequoiataxrelief.com" `
Connect-PnPOnline
                Set-PnPList -Identity "Documents" -ForceCheckout $false

                # Modify navigation
                Show-StatusWindow "Creating quicklaunch..."
                $context = Get-PnPContext
                $web = $context.Web
                $context.Load($web)
                $context.ExecuteQuery()
                Write-Log "Got PnP context. Creating quick launch..."
                $quickLaunch = $web.Navigation.QuickLaunch
                $context.Load($quickLaunch)
                $context.ExecuteQuery()
                Write-Log "Editing quicklaunch..."
                
                # Remove all existing navigation nodes
                Show-StatusWindow "Removing existing navigation nodes..."
                $nodesToRemove = @()
                foreach ($node in $quickLaunch){
                    $nodesToRemove += $node
                }
                $totalNodes = $nodesToRemove.Count
                $countNode = 0
                foreach ($node in $nodesToRemove) {
                    $countNode++
                    $node.DeleteObject()
                    Write-Log "Removed node $countNode / $totalNodes : $($node.Title)"
                }
                $context.ExecuteQuery()
                Write-Log "Removed all $totalNodes nodes successfully"
                
                # Create new navigation nodes
                Show-StatusWindow "Creating new navigation nodes..."
                Write-Log "Creating node for Documents folder"
                $documentsUrl = $web.ServerRelativeUrl + "/Documents"
                $nodeInfo = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
                $nodeInfo.Title = "Documents"
                $nodeInfo.Url = $documentsUrl
                $nodeInfo.AsLastNode = $true
                $quickLaunch.Add($nodeInfo)
                $context.ExecuteQuery()
                Write-Log "Documents added to QuickLaunch"

                # Create Client Documents folder
                Show-StatusWindow "Creating Client Documents folder..."
                Write-Log "Creating Documents/Client Documents"
                Add-PnPFolder -Name "Client Documents" -Folder "Documents"
                Write-Log "Creating share link for client email address"
                Show-StatusWindow "Sharing Client Documents folder..."
                $sharingResult = Add-PnPFolderUserSharingLink -Folder "Documents/Client Documents" `
                    -ShareType Edit -Users $clientEmail
                Write-Log "Shared with: $clientEmail, with sharing link:$($sharingResult.WebUrl)"

                # Set up alerts if not skipped
                if (-not $skipAlerts) {
                    Show-StatusWindow "Creating alerts..."
                    $documentLibrary = Get-PnPList -Identity "Documents"
                    $group = Get-PnPGroup -Identity $TeamName
                    $users = Get-PnPGroupMember -Identity $group

                    foreach ($user in $users) {
                        Add-PnPAlert -List $documentLibrary -User $user.LoginName `
                            -Title "New Document Alert for $TeamName" `
                            -DeliveryMethod Email -ChangeType All -Frequency Immediate
                        Write-Log "Created Alert for $($user.LoginName)"
                    }
                    Write-Log "Alerts completed"
                } else{
                    Write-Log "Alerts skipped"
                }

                # Determine the From email address
                $currentUserEmail = if ($From) { $From } else { "$env:USERNAME@sequoiataxrelief.com" }
                Show-StatusWindow "Preparing emails..."
                $companyLogoUrl = Get-CompanyUrl -UrlType "Logo"
                $supportSiteUrl = Get-CompanyUrl -UrlType "Support"
                $privacyPolicyUrl = Get-CompanyUrl -UrlType "PrivacyPolicy"

                # Prepare placeholders
                $placeholders = @{
                    ClientName = $clientName
                    CompanyLogoUrl = $companyLogoUrl
                    ClientSharingLink = $sharingResult.WebUrl
                    SupportSiteUrl = $supportSiteUrl
                    PrivacyPolicyUrl = $privacyPolicyUrl
                    CurrentYear = (Get-Date).Year
                    TeamName = $TeamName
                }
                Write-Log "Placeholders: $(ConvertTo-CompressedJson $placeholders)"
                
                if (!$skipEmails) {
                    # Send client email if not switch skipClientEMail
                    try {
                        Show-StatusWindow "Getting client email template..."
                        Write-Log "Attempting to send client email"
                        $clientEmailTemplate = Get-EmailTemplate -TemplateName $EmailTemplateName
                        Write-Log "Client email template retrieved"
                        $clientEmailBody = Format-EmailTemplate -Template $clientEmailTemplate -Placeholders $placeholders
                        Write-Log "Client email body formatted"
                        Show-StatusWindow "Sending client email..."
                        Send-PnPMailWithTimeout -From $currentUserEmail -To $clientEmail `
                            -Subject "Welcome to Your Secure Client Portal - Sequoia Tax Relief" `
                            -Body $clientEmailBody -BodyContentType Html
                        Write-Log "Client email sent successfully to $clientEmail"
                    } catch {
                        Write-Log "Failed to send client email to $clientEmail. Error: $_" -Level "ERROR"
                    }

                    Write-Log "Starting team email process"

                    # Send team email
                    try {
                        Show-StatusWindow "Getting team email template..."
                        Write-Log "Attempting to send team email"
                        $teamEmailTemplate = Get-EmailTemplate -TemplateName "FullClientTeam"
                        Write-Log "Team email template retrieved"
                        $teamEmailBody = Format-EmailTemplate -Template $teamEmailTemplate -Placeholders $placeholders
                        Write-Log "Team email body formatted"
                        Show-StatusWindow "Getting team users..."
                        Write-Log "Connecting to PnP for Team Members email"
                        Connect-PnPOnline -Url "https://sequoiataxrelief.sharepoint.com/sites/secureclientupload" `
                            -ClientId 66169a0f-1b2f-49af-acba-024033664ec1 `
                            -Tenant "sequoiataxrelief.com" `
Connect-PnPOnline
                        Write-Log "Connected, Retrieving team members for $TeamName"
                        $teamUsers = Get-PnPGroupMember -Identity $TeamName | Select-Object -ExpandProperty Email
                        Write-Log "Team members retrieved: $($teamUsers -join ', ')"
                        
                        $teamEmailRecipients = @($currentUserEmail) + $teamUsers
                        # Clean the email list of any empty or whitespace emails
                        $teamEmailRecipients = $teamEmailRecipients | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
                        Write-Log "Sending team email to: $($teamEmailRecipients -join ', ')"
                        Show-StatusWindow "Sending team email..."
                        Send-PnPMailWithTimeout -From $currentUserEmail -To $teamEmailRecipients `
                            -Subject "New Full Client Subsite Created - Action Required" `
                            -Body $teamEmailBody -BodyContentType Html
                        Write-Log "Team email sent successfully"
                    } catch {
                        Write-Log "Failed to send team email. Error: $_" -Level "ERROR"
                    }

                    Write-Log "Team email process completed"
                    Show-StatusWindow "Client/Team sent successfully"
                } else {
                    Write-Log "Emails Skipped, would have sent:"
                    Write-Log "Send-PnPMailWithTimeout -From $currentUserEmail -To $clientEmail -Subject 'Welcome to Your Secure Client Portal - Sequoia Tax Relief' -Body $clientEmailBody -BodyContentType Html"
                    Write-Log "Send-PnPMailWithTimeout -From $currentUserEmail -To $teamEmailRecipients -Subject 'New Full Client Subsite Created - Action Required' -Body $teamEmailBody -BodyContentType Html"
                }
                # Send payments email if checkbox is checked
                if ($sendPaymentsEmail) {
                    try {
                        Show-StatusWindow "Getting payments email template..."
                        Write-Log "Attempting to send payments email"
                        $paymentsEmailTemplate = Get-EmailTemplate -TemplateName "FullClientPayments"
                        Write-Log "Payments email template retrieved"
                        $paymentsEmailBody = Format-EmailTemplate -Template $paymentsEmailTemplate -Placeholders $placeholders
                        Write-Log "Payments email body formatted"
                        Show-StatusWindow "Sending payments email..."
                        Send-PnPMailWithTimeout -From $currentUserEmail -To "payments@sequoiataxrelief.com" `
                            -Subject "New Full Client Subsite Created - Payment Processing Required" `
                            -Body $paymentsEmailBody -BodyContentType Html
                        Write-Log "Payments email sent successfully"
                    } catch {
                        Write-Log "Failed to send payments email. Error: $_" -Level "ERROR"
                    }
                }
                # Update JSON cache
                try {
                    Show-StatusWindow "Updating JSON cache..."
                    Update-ClientJsonCache -teamName $TeamName -clientName $clientName -email $clientEmail `
                        -sharingLink $sharingResult.WebUrl -isNewSubsite
                    Write-Log "JSON cache updated successfully"
                } catch {
                    Write-Log "Failed to update JSON cache: $_" -Level "ERROR"
                }

                Show-StatusWindow "Subsite created successfully!"
                Start-Sleep -Seconds 2
                Remove-PnPConnection
                Show-StatusWindow -Close
                Write-Log "Email sending process completed"
                return "Success"
            } else {
                Show-StatusWindow "Error: Subsite already exists"
                Start-Sleep -Seconds 2
                Remove-PnPConnection
                Show-StatusWindow -Close
                Write-Log "Subsite already exists: $fullSubsiteUrl"
                return "Error: Subsite already exists"
            }
        }
        catch {
            Write-Log "Error in New-ClientSubsite: $($_.Exception.Message)" -Level "ERROR"
            Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
            Show-StatusWindow "Error: $($_.Exception.Message)"
            Remove-PnPConnection
            Start-Sleep -Seconds 5
            Show-StatusWindow -Close
            return "Error: $($_.Exception.Message)"
        }
    }
    end {
        # Ensure connections are closed
        Show-StatusWindow -Close
        Remove-PnPConnection
    }
}
