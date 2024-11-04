using namespace System.Windows.Forms
using namespace System.Drawing

function Invoke-Cleanup {
    param (
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.Form]$Form
    )
    Write-Log "Helper function preforming clean up..."
    Import-Module PnP.PowerShell
    
    foreach ($cleanupItem in $script:itemsToCleanup) {
        switch ($cleanupItem.Type) {
            "Job" {
                try {
                    Remove-Job -Id $cleanupItem.Item -Force -ErrorAction SilentlyContinue
                    Write-Log "Removed job with ID: $($cleanupItem.Item)"
                } catch {
                    Write-Log "Error removing job with ID $($cleanupItem.Item): $_" -Level "ERROR"
                }
            }
            "Process" {
                try {
                    $process = $cleanupItem.Item
                    if (!$process.HasExited) {
                        $process.CloseMainWindow()
                        if (!$process.WaitForExit(5000)) {
                            $process.Kill()
                        }
                    }
                    Write-Log "Closed child process: $($process.Id)"
                } catch {
                    Write-Log "Error closing child process $($process.Id): $_" -Level "ERROR"
                }
            }
        }
    }
    $script:itemsToCleanup = @()
    # Disconnect from PnP
    Remove-PnPConnection
    Write-Log "Helper function cleanup complete"
}
