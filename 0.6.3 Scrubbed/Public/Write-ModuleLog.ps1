using namespace System.Windows.Forms
using namespace System.Drawing

function Write-ModuleLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    
    # Call the private Write-Log function
    $private:WriteLog = Get-Command -Module SequoiaTax -Name 'Write-Log' -CommandType Function
    & $private:WriteLog -Message $Message -Level $Level
}
