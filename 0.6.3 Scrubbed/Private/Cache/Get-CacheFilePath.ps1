using namespace System.Windows.Forms
using namespace System.Drawing

function Get-CacheFilePath { # gets the path for the cache file based on the cache type
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Folder", "SubsiteTeamAmanda", "SubsiteTeamJamie", "SubsiteTeamKyle", "Redwood")]
        [string]$CacheType
    )
    Write-Log "Getting cache file path for type: $CacheType"
    #$basePath = "\\str-0111\MainMenu\0.6.3-Local\cache"
    Write-Log "Base path: $script:cachePath"
    switch ($CacheType) {
        "Folder" { 
            $fullPath = Join-Path $script:cachePath "client_folder_cache.json"
            Write-Log "Returning path for Folder: $fullPath"
            return $fullPath
        }
        "SubsiteTeamAmanda" { 
            $fullPath = Join-Path $script:cachePath "client_subsite_cache_TeamAmanda.json"
            Write-log "Returning path for TeamAmanda: $fullPath"
            return $fullPath
        }
        "SubsiteTeamJamie" { 
            $fullPath = Join-Path $script:cachePath "client_subsite_cache_TeamJamie.json"
            Write-Log "Returning path for TeamJamie: $fullPath"
            return $fullPath
        }
        "SubsiteTeamKyle" { 
            $fullPath = Join-Path $script:cachePath "client_subsite_cache_TeamKyle.json"
            Write-Log "Returning path for TeamKyle: $fullPath"
            return $fullPath
        }   
        "Redwood"{
            $fullPath = Join-Path $script:cachePath "redwood_client_cache.json"
            Write-Log "Returning path for Redwood: $fullPath"
            return $fullPath
        }
    }
}
