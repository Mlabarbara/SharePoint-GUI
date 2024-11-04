using namespace System.Windows.Forms
using namespace System.Drawing

function Remove-CleanupItem {
    param(
        [Parameter(Mandatory=$true)]
        [PSObject]$Item,
        [Parameter(Mandatory=$true)]
        [string]$Type
    )
    $initialCount = $script:itemsToCleanup.Count
    $script:itemsToCleanup = $script:itemsToCleanup | Where-Object { $_.Item -ne $Item -or $_.Type -ne $Type }
    $removedCount = $initialCount - $script:itemsToCleanup.Count
    if ($removedCount -gt 0) {
        Write-Log "Successfully removed $removedCount cleanup item(s)"
    } else {
        Write-Log "No matching cleanup items found to remove"
    }
}
