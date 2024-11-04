# Private/SharePoint/Invoke-PnPOperationAsync.ps1
function Invoke-PnPOperationAsync {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$Parameters = @{},
        
        [Parameter(Mandatory=$false)]
        [int]$TimeoutSeconds = 300,
        
        [Parameter(Mandatory=$false)]
        [switch]$UseRunspace
    )
    
    try {
        if ($UseRunspace) {
            # Create and configure runspace
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
            $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $runspace.Open()
            
            # Initialize runspace with required modules
            $runspace.SessionStateProxy.SetVariable('logFile', $global:logFile)
            $runspace.SessionStateProxy.ImportPSModule('PnP.PowerShell')
            $runspace.SessionStateProxy.ImportPSModule('SequoiaTax')
            
            $powershell = [powershell]::Create().AddScript($ScriptBlock)
            if ($Parameters.Count -gt 0) {
                $powershell.AddParameters($Parameters)
            }
            $powershell.Runspace = $runspace
            
            Write-Log "Starting async operation in runspace"
            $asyncResult = $powershell.BeginInvoke()
            
            # Wait for completion or timeout
            if (-not $asyncResult.AsyncWaitHandle.WaitOne($TimeoutSeconds * 1000)) {
                throw "Operation timed out after $TimeoutSeconds seconds"
            }
            
            $result = $powershell.EndInvoke($asyncResult)
            return $result
        }
        else {
            # Use background job
            $job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $Parameters
            
            Write-Log "Starting async operation as job"
            $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
            
            if (-not $completed) {
                Remove-Job -Job $job -Force
                throw "Operation timed out after $TimeoutSeconds seconds"
            }
            
            $result = Receive-Job -Job $job
            Remove-Job -Job $job
            return $result
        }
    }
    catch {
        Write-Log "Error in async operation: $_" -Level "ERROR"
        throw
    }
    finally {
        if ($UseRunspace) {
            if ($powershell) { $powershell.Dispose() }
            if ($runspace) { $runspace.Dispose() }
        }
    }
}
