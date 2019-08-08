# scripts

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
1. and the ablitiy to apply the optional filters for the non-destructive, listing runs as well.

### sync_odrive.ps1
Sync all your cloud files using [odrive cli](https://docs.odrive.com/docs/odrive-cli), instead of syncing them one-by-one. File syncing is parallelised (use the `-JobCount` parameter) and there's a progress bar too.
