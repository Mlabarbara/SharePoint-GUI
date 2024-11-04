function Get-EmailSentVerification {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [Parameter(Mandatory=$true)]
        [string[]]$Recipients, 
        [Parameter(Mandatory=$true)]
        [int]$WaitSeconds,
        [Parameter(Mandatory=$false)]
        [string]$LogFile
    )
    # Start a new job for the verification to avoid DLL conflicts
    $verificationJob = Start-Job -ScriptBlock {
        param($Subject, $Recipients, $WaitSeconds, $LogFile)
        
        function Write-VerificationLog {
            param([string]$Message, [string]$Level = "INFO")
            if ($LogFile) {
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $logMessage = "$timestamp [$Level] [Verify-EmailSent] $contextInfo $Message"
                Add-Content -Path $LogFile -Value $logMessage
            }
        }

        function Test-EmailExists {
            param($UserId, $Subject, $Recipients)
            
            $messages = Get-MgUserMessage -UserId $UserId -Top 10
            
            foreach ($message in $messages) {
                # Check subject match
                if ($message.Subject -eq $Subject) {
                    # Get recipient addresses
                    $messageRecipients = $message.ToRecipients.EmailAddress | 
                        Select-Object -ExpandProperty Address
                    
                    # Check if any recipient matches
                    $recipientMatch = $Recipients | Where-Object { 
                        $recipient = $_
                        $messageRecipients -contains $recipient
                    }
                    
                    if ($recipientMatch) {
                        Write-VerificationLog "Found matching email - Subject: $Subject, Recipient: $($recipientMatch)"
                        return $true
                    }
                }
            }
            return $false
        }

        try {
            Import-Module Microsoft.Graph.Mail
            Connect-MgGraph -ClientId "a040682b-2c17-4d61-93c6-2e1797a22f08" -TenantId "036013a7-97b0-44e1-ba2f-8bdbd8432ec3" -Certificate "\\str-0111\MainMenu\msGraphcert\MgCertWithPrivate.pfx" -NoWelcome | Out-Null
            
            Write-VerificationLog "Starting verification for subject: $Subject"
            
            # Get user ID first
            $userId = Get-MgUser | 
                Where-Object { $_.Mail -like "$env:USERNAME@*" } |
                Select-Object -First 1 -ExpandProperty Id

            if (!$userId) {
                Write-VerificationLog "Could not find user ID" -Level "ERROR"
                return @{
                    Status = "Error"
                    Message = "Could not find user ID"
                    Found = $false
                } | ConvertTo-Json
            }

            # Check immediately first
            if (Test-EmailExists -UserId $userId -Subject $Subject -Recipients $Recipients) {
                return @{
                    Status = "Success"
                    Message = "Email verified immediately"
                    Found = $true
                } | ConvertTo-Json
            }

            # Exponential backoff checks
            $delays = @(1, 2, 4)
            foreach ($delay in $delays) {
                Write-VerificationLog "Email not found, waiting $delay seconds before next check..."
                Start-Sleep -Seconds $delay
                
                if (Test-EmailExists -UserId $userId -Subject $Subject -Recipients $Recipients) {
                    return @{
                        Status = "Success"
                        Message = "Email verified after $delay second delay"
                        Found = $true
                    } | ConvertTo-Json
                }
            }

            Write-VerificationLog "Email not found after all attempts" -Level "WARN"
            return @{
                Status = "Warning"
                Message = "Email not found after all verification attempts"
                Found = $false
            } | ConvertTo-Json
        }
        catch {
            Write-VerificationLog "Error verifying email: $_" -Level "ERROR"
            return @{
                Status = "Error"
                Message = $_.ToString()
                Found = $false
            } | ConvertTo-Json
        }
        finally {
            Disconnect-MgGraph | Out-Null
        }
    } -ArgumentList $Subject, $Recipients, $WaitSeconds, $LogFile

    # Wait for verification job to complete
    $completed = Wait-Job -Job $verificationJob -Timeout 30
    if (-not $completed) {
        Stop-Job -Job $verificationJob
        Remove-Job -Job $verificationJob
        return @{
            Status = "Error"
            Message = "Verification timed out after 30 seconds"
            Found = $false
        } | ConvertTo-Json
    }

    $result = Receive-Job -Job $verificationJob
    Remove-Job -Job $verificationJob
    return $result
}
