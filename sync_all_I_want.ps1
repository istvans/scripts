. .\utils.ps1
.\sync_odrive.ps1 -ExcludePattern "Shared With Me|Steve-OneDrive\\Videos\\|Steve-OneDrive\\Pictures\\|Amazon Cloud Drive" -FolderExpandExcludePattern "Shared With Me" -JobCount 16
pause "Press any key to continue"
