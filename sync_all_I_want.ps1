. $PSScriptRoot\utils.ps1
& $PSScriptRoot\sync_odrive.ps1 -ExcludePattern "Shared With Me|Steve-OneDrive\\Videos\\|Steve-OneDrive\\Pictures\\|Amazon Cloud Drive" -FolderExpandExcludePattern "Shared With Me" -JobCount 12
pause "Press any key to continue"
