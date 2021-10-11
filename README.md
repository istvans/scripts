# scripts

## Cross-platform

### checksum_file.py
Print the checksum of a file, calculated using common, standard algorithms.

### expand_every_cloudf.py
Traverse a directory recursively until no cloudf placeholders left.
(This might only work on Windows. TODO try on Linux.)

## Linux

### build_kodi_waipu_addon_on_rpi_osmc.sh
An automated way of building the [Waipu.tv](https://github.com/flubshi/pvr.waipu) PVR plugin for Kodi on a Raspberry Pi OSMC.
By default it builds the addon for Kodi Leia, but another version can be specified as an argument too.

## Windows

### removeGhosts.ps1
I found the original version of this script [here](https://theorypc.ca/2017/06/28/remove-ghost-devices-natively-with-powershell/?unapproved=10049&moderation-hash=10c8951dd7472325cbcaeed99af2ec9e).
Thank you for TrententTye and Alexander Boersch for their excellent work!

In this repo I publish a slightly improved version of this script which adds:
1. -narrowByClass,
1. -narrowByFriendlyName,
1. -force, making the default run asking for confirmation,
1. and the ability to apply the optional filters for the non-destructive, listing runs as well.

### sync_odrive.ps1
Sync all your cloud files using [odrive cli](https://docs.odrive.com/docs/odrive-cli), instead of syncing them one-by-one. File syncing is parallelised (use the `-JobCount` parameter) and there's a progress bar too.

It uses `expand_every_cloudf.py` to expand every folder before it starts the syncing. This way the syncing can run once without missing any file.

### phone_vs_cloud.ps1
Thank you for Daiyan Yingyu for publishing the [original version](https://blog.daiyanyingyu.uk/files/MoveFromPhone.ps1) of the script.

This script scans a directory on a phone and checks whether every file can be found in another directory or any of its subdirectories (recursively). If a file is missing, it gets copied over to a specified directory, which may or may not be the same as where the files are looked for. The script finishes with a summary of how many files were copied from how many it found.

This is handy if you have a local sync of your cloud drive and want to make sure every file is backed up before freeing up space on your device. The OneDrive android app has a feature that can remove the uploaded files, but that feature is missing on Android 11 (as listed [here](https://support.microsoft.com/en-us/office/fixes-or-workarounds-for-recent-issues-in-onedrive-36110213-f3f6-490d-8cb7-3833539def0b)). I couldn't really find any other *reliable* way to know that all the files have been indeed uploaded.
