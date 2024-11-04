[CmdletBinding(SupportsShouldProcess)]
param()

$folderPath = "X:\Sharepoint\Modules\SequoiaTax\0.6.3 Scrubbed"

Get-ChildItem -Path $folderPath -Recurse -Include *.ps1,*.psm1,*.psd1 | ForEach-Object {
    $file = $_
    # Read the content as an array of lines to track line numbers
    $lines = Get-Content $file.FullName

    for ($i = 0; $i -lt $lines.Count; $i++) {
Connect-PnPOnline
            $originalLine = $lines[$i]
            
            # Handle multi-line entries with backticks
            while ($originalLine.Trim().EndsWith('`') -and ($i + 1 -lt $lines.Count)) {
                $i++
                $originalLine += "`n" + $lines[$i]
            }

            if ($PSCmdlet.ShouldProcess(
                "File: $($file.FullName)",
Connect-PnPOnline
            )) {
Connect-PnPOnline
                # Clear any continuation lines
                while ($originalLine.Trim().EndsWith('`') -and ($i + 1 -lt $lines.Count)) {
                    $i++
                    $lines[$i] = ""
                }
            }
        }
    }

    if ($PSCmdlet.ShouldProcess($file.FullName, "Save changes")) {
        $lines | Set-Content $file.FullName -Force
    }
}
