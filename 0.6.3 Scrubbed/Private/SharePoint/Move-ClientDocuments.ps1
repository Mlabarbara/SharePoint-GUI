using namespace System.Windows.Forms
using namespace System.Drawing

function Move-ClientDocuments{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$clientName,
        [Parameter(Mandatory=$true)]
        [string]$TeamName,
        [Parameter(Mandatory=$true)]
        [string]$selectedClientUrl
    )
    begin {
        Import-Module PnP.PowerShell
        Import-Module "$script:LOCAL_ROOT\HelperFunctions.psm1" -Force
        Import-Module "$script:LOCAL_ROOT\SharePointModule.psm1" -Force
        function Remove-EmptyFolder {
            param ([string]$FolderPath)
            Write-Log "removing Folder $FolderPath"
            $maxRetries = 3
            $retryCount = 0
            Show-StatusWindow "Removing folder from Shared Documents"
            while ($retryCount -lt $maxRetries) {
                $remainingItems = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderPath -Recursive
                if ($remainingItems.Count -eq 0) {
                    try {
                        $folderToDelete = $FolderPath.Split('/')[-1]
                        Write-Log "Empty folder found: $folderToDelete"
                        $parentFolder = $FolderPath -replace "/[^/]+$"
                        Remove-PnPFolder -Name $folderToDelete -Folder $parentFolder -Force
                        Write-Log "Empty folder Deleted: $folderToDelete"
                        return @{
                            Status = "Success"
                            Message = "Empty folder Deleted: $folderToDelete"
                        }
                    }
                    catch {
                        Write-Log "Error removing folder: $($_.Exception.Message)" -Level "ERROR"
                    }
                }
                else {
                $retryCount++
                if ($retryCount -lt $maxRetries) {
                    Write-Log "Folder not empty, retrying in 500ms" -Level "WARN"
                    Start-Sleep -Milliseconds 500
                    }
                }
            }
            Show-StatusWindow -Close
            $finalErrorMessage = "Failed to remove empty folder after $maxRetries attempts"
            Write-Log $finalMessage -Level "ERROR"
            return @{
                Status = "Error"
                Message = $finalErrorMessage
            }
        }
        function Test-NewSubsite {
            param ([string]$FullSubsiteUrl)
            Show-StatusWindow "Checking if new subsites' Client Documents exists..."
Connect-PnPOnline
            Write-Log "Connected to new subsite: $FullSubsiteUrl"
        
            $clientSubsiteFolderPath = "Documents/Client Documents"
            if (-not (Get-PnPFolder -Url $clientSubsiteFolderPath)) {
                Write-Log "Folder does not exist, Creating Client Documents folder. "
                New-PnPFolder -Name "Client Documents" -Folder "Documents"
                Write-Log "Created Client Documents folder"
            }
            Show-StatusWindow -Close
            return Get-PnPFolder $clientSubsiteFolderPath
        }
        function Move-Items {
            param (
                $Items,
                $TargetFolder,
                [string]$SourcePath
            )
            $files = $Items | Where-Object { $_.GetType().Name -eq "File" }
            $folders = $Items | Where-Object { $_.GetType().Name -eq "Folder" }
            Write-Log "Found $($files.Count) files and $($folders.Count) folders to move"
            Show-StatusWindow "Moving Folders to Subsite, large files or many files may take a while..."
            foreach ($folder in $folders) {
                Write-Log "Moving $($folder.Name) to $($TargetFolder.ServerRelativeUrl)" -Level "INFO"
                Move-PnPFile -SourceUrl $folder.ServerRelativeUrl -TargetUrl $TargetFolder.ServerRelativeUrl -Force
                Start-Sleep -Milliseconds 500
            }
            Show-StatusWindow "Moving Files to Subsite, large files or many files may take a while..."
            foreach ($file in $files) {
                Write-Log "Moving $($file.Name) to $($TargetFolder.ServerRelativeUrl)" -Level "INFO"
                if ($file.ServerRelativeUrl -notmatch [regex]::Escape($SourcePath + "/[^/]+/")) {
                    Move-PnPFile -SourceUrl $file.ServerRelativeUrl -TargetUrl $TargetFolder.ServerRelativeUrl -Force -NoWait
                    Start-Sleep -Milliseconds 500
                }
            }
            Show-StatusWindow -Close
            return @{
                Status = "Success"
                MovedFiles = $files.Count
                MovedFolders = $folders.Count
            }
        }
    }
    process {
        Write-Log "Starting Move-ClientDocuments function"
        Write-Log "Parameters: clientName=$clientName, TeamName=$TeamName, selectedClientUrl=$selectedClientUrl"
        try {
            Show-StatusWindow "Getting client documents from folder..."
            # move documents to the new subsite
            $originalSiteUrl = "https://sequoiataxrelief.sharepoint.com/sites/secureclientupload"
            $siteUrlBase = Get-SharePointBaseUrl
            $siteUrl = "$siteUrlBase/$TeamName"
            $sanitizedClientName = $clientName -replace '\s', '' -replace '[^a-zA-Z0-9]', ''
            $fullSubsiteUrl = "$siteUrl/$sanitizedClientName"
            $documentsPath = $selectedClientUrl -replace "^.*?(Shared Documents.*)", '$1'
            Write-Log "Moving documents to new subsite: $fullSubsiteUrl"
            
            #connect to the origional site
Connect-PnPOnline
            Write-Log "Connected to original site: $originalSiteUrl"
            # get the relative path of the selected folder
            
            $items = Get-PnPFolderItem -FolderSiteRelativeUrl $documentsPath 

            if ($items.Count -eq 0) {
                Write-Log "No items found in the folder. Skipping move operation."
                $removeResult = Remove-EmptyFolder -FolderPath $documentsPath
                return @{
                    Status = "EmptyFolder"
                    Message = $removeResult.Message
                    FolderRemovalStatus = $removeResult.Status
                }
            }
            Show-StatusWindow -Close
            # check new subsite and move items
            Write-Log "Running Test-NewSubsite"
            $newSubsiteFolder = Test-NewSubsite -FullSubsiteUrl $fullSubsiteUrl
            Write-Log "Running Move-Items"
            $moveResult = Move-Items -Items $items -TargetFolder $newSubsiteFolder -SourceFolder $documentsPath

            # cleanup origional folder
Connect-PnPOnline
            Write-Log "Running Remove-Empty Folder"
            $removeResult = Remove-EmptyFolder -FolderPath $documentsPath

            if ($removeResult.Status -eq "Error") {
                Write-Log "Error removing empty folder: $($removeResult.Message)" -Level "ERROR"
            }

            return @{
                Status = "Success"
                MovedFiles = $moveResult.MovedFiles
                MovedFolders = $moveResult.MovedFolders
                FolderRemovalStatus = $removeResult.Status
                FolderRemovalMessage = $removeResult.Message
            }
        }
        catch {
            Write-Log "Error in Move-ClientDocuments: $($_.Exception.Message)" -Level "ERROR"
            return @{ 
                Status = "Error" 
                Message = $($_.Exception.Message) 
            }
        }
         finally {
            Show-StatusWindow -Close
        }
    }
}
