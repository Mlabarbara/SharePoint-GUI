using namespace System.Windows.Forms
using namespace System.Drawing

function Remove-PnPConnection { # TO USE: Remove-PnPConnection
    try {
        Write-Log "Attempting to remove PnPOnline connection"
        $pnpClosingCheck = Get-PnPConnection -ErrorAction Stop
        if ($pnpClosingCheck) {
            Write-Log "Disconnecting from PnPOnline"
Connect-PnPOnline
            Write-Log "Disconnected from PnPOnline"
        }
    }
    catch {
        if ($_.Exception.Message -like "*The current connection holds no SharePoint context*") {
            Write-Log "No active PnPOnline connection to disconnect" -Level "INFO"
        }
        else {
            Write-Log "Error disconnecting from PnPOnline: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}
