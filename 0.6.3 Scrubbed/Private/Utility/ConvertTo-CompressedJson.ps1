using namespace System.Windows.Forms
using namespace System.Drawing

function ConvertTo-CompressedJson {
    param([Parameter(Mandatory=$true, ValueFromPipeline=$true)] $InputObject)
    
    $jsonString = $InputObject | ConvertTo-Json -Depth 10 -Compress
    return $jsonString
}
