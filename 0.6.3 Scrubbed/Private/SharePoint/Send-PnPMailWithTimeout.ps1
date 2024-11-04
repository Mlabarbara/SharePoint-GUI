using namespace System.Windows.Forms
using namespace System.Drawing

function Send-PnPMailWithTimeout { # TO USE: Send-PnPMailWithTimeout -From $currentUserEmail -To $clientEmail -Subject "Welcome to Your Secure Client Portal - Sequoia Tax Relief" -Body $clientEmailBody -BodyContentType Html
param(
        [string]$From,
        [string[]]$To,
        [string]$Subject,
        [string]$Body,
        [string]$BodyContentType = "Html",
        [int]$TimeoutSeconds = 60,
        [int]$VerificationTimeOut = 5
    )
    Write-Log "Attempting to send email. From: $From, To: $To, Subject: $Subject, Body length: $($Body.Length)"
    
    if ([string]::IsNullOrWhiteSpace($Body)) {
        Write-Log "Error: Email body is empty or whitespace" -Level "ERROR"
        return @{
            Status = "Error"
            To = $To 
            Message = "Email body is empty or whitespace, did not send."
        }
    }

    # We can use Invoke-PnPOperationAsync for the email sending
    try {
        $emailParams = @{
            From = $From
            To = $To
            Subject = $Subject
            Body = $Body
            BodyContentType = $BodyContentType
        }

        $result = Invoke-PnPOperationAsync -ScriptBlock {
            param($params)
            Connect-PnPOnline -Url "https://sequoiataxrelief.sharepoint.com/sites/secureclientupload" `
                -ClientId 66169a0f-1b2f-49af-acba-024033664ec1 `
                -Tenant "sequoiataxrelief.com" `
Connect-PnPOnline
            
            Send-PnPMail @params
Connect-PnPOnline
            
            return @{
                Status = "Success"
                To = $params.To
                Message = "Email sent successfully"
            }
        } -Parameters $emailParams -TimeoutSeconds $TimeoutSeconds

        if ($result.Status -eq "Success") {
            Write-Log "Email sent, starting verification..."
            $verificationResult = Get-EmailSentVerification -Subject $Subject -Recipients $To `
                -WaitSeconds $VerificationTimeOut -LogFile $script:LOGFILE | ConvertFrom-Json

            return @{
                Status = "Success"
                To = $To
                Message = "Email sent"
                Verified = $verificationResult.Found
                VerificationMessage = $verificationResult.Message
            }
        }
    }
    catch {
        Write-Log "Error sending email: $_" -Level "ERROR"
        return @{
            Status = "Error"
            To = $To
            Message = "Failed to send email: $_"
            Verified = $false
        }
    }
}
