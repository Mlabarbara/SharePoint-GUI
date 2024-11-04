using namespace System.Windows.Forms
using namespace System.Drawing

function Remove-SpacesAndPunc {
    param([string]$name)
    return $name -replace '[^\w]', ''
}
