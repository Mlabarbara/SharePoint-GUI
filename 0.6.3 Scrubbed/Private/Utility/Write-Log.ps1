using namespace System.Windows.Forms
using namespace System.Drawing

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO",

        [Parameter(Mandatory=$false)]
        [string]$LogFile
    )
    if ([string]::IsNullOrWhiteSpace($Message)){
        return
    }
    # Use whichever log file variable is available
    $logFilePath = if ($LogFile) { 
        $LogFile
    } elseif ($global:LOGFILE) { 
        $global:LOGFILE 
    } else {
        # Default to network path with version
        Join-Path $script:ModuleLogPath "$($env:USERNAME)--$script:ModuleVersion--$((Get-Date).ToString('M-dd-yy')).log"
    }
    $callStack = Get-PSCallStack
    $callerInfo = $callStack[1]  # Index 1 is the immediate caller of Write-Log
    $scriptName = Split-Path -Leaf $callerInfo.ScriptName
    $functionName = $callerInfo.FunctionName
    $lineNumber = $callerInfo.ScriptLineNumber
    
    if ($functionName -eq "<ScriptBlock>") {
        $contextInfo = "[$scriptName::$functionName::$lineNumber]"
    } elseif (![string]::IsNullOrEmpty($functionName)) {
        $contextInfo = "[$scriptName::$functionName::$lineNumber]"
    } else {
        $contextInfo = "[$scriptName::$lineNumber]"
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $contextInfo $Message"
    Add-Content -Path $logFilePath -Value $logMessage
}
