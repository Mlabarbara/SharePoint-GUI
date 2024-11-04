using namespace System.Windows.Forms
using namespace System.Drawing

function Add-CleanupItem { # TO USE: Add-CleanupItem  -Item $job.Id -Type "Job | Process"
    param(
        [Parameter(Mandatory=$true)]
        [PSObject]$Item,
        [Parameter(Mandatory=$true)]
        [string]$Type
    )
    $script:itemsToCleanup += @{Item = $Item; Type = $Type}
}
