# Ensure all the files are uploaded from the phone to the cloud.
# Copy any missing files to a destination folder which should be a cloud sync folder.
param(
    [Parameter(Mandatory)]
    [string]$phoneName,
    [string]$phoneFolderPath="<configured>",
    [string]$cloudFolderPath="<configured>",
    [string]$destinationFolderPath="<configured>",
    [string]$filter=".(jpg|mp4)$",
    [Parameter(HelpMessage="See cloud_sync_config.example.ps1.")]
    [string]$configFile="cloud_sync_config.ps1",
    # TODO add multiprocessing [uint32]$jobCount=$(Get-ComputerInfo -Property CsProcessors).CsProcessors.NumberOfCores,
    [switch]$confirmCopy=$false,
    [switch]$dryRun=$false
)

function Get-ShellProxy
{
    if (-not $global:ShellProxy)
    {
        $global:ShellProxy = new-object -com Shell.Application
    }
    $global:ShellProxy
}

function Get-Phone
{
    param($phoneName)
    $shell = Get-ShellProxy
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
    # => "My Computer" â€” the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
    # This folder can also contain mapped network drives.
    $shellItem = $shell.NameSpace(17).self
    $phone = $shellItem.GetFolder.items() | where { $_.name -eq $phoneName }
    return $phone
}

function Get-SubFolder
{
    param($parent,[string]$path)
    $pathParts = @( $path.Split([system.io.path]::DirectorySeparatorChar) )
    $current = $parent
    foreach ($pathPart in $pathParts)
    {
        if ($pathPart)
        {
            $current = $current.GetFolder.items() | where { $_.Name -eq $pathPart }
        }
    }
    return $current
}

<#
.SYNOPSIS
Return a possible equivalent of the name of `$phoneFile` in the cloud when
"Keep file names as in the device" is disabled.
#>
function GetMegaFilename
{
    param($phoneFile)

    $extension = [System.IO.Path]::GetExtension($phoneFile.name)
    $megaName = $phoneFile.ExtendedProperty("System.DateModified").ToString("yyyy-MM-dd HH.mm.ss")
    $megaFilename = "$megaName$extension"
    return $megaFilename
}

function ByteToMegabyte
{
    param($byteSize)
    return [math]::floor($byteSize / 1MB)
}

function IsInCloud
{
    param($phoneFile, $cloudFiles)

    $originalFilename = $phoneFile.name
    $cloudFile = $cloudFiles[$originalFilename]

    if ($cloudFile -eq $null) {
        $megaFilename = GetMegaFilename($phoneFile)
        Write-Debug "... Mega filename guess for '$originalFilename': '$megaFilename'"
        $cloudFile = $cloudFiles[$megaFilename]
        Write-Debug "... Match: $($cloudFile -ne $null)"
    }

    if ($cloudFile -eq $null) {
        $result = $false
    } else {
        # TODO winmerge? Get-FileHash?
        $phoneFileSize = ByteToMegabyte($phoneFile.ExtendedProperty("System.Size"))
        $cloudFileSize = ByteToMegabyte($cloudFile.length)
        if ($phoneFileSize -eq $cloudFileSize) {
            $result = $true
        } else {
            Write-Debug "... $($phoneFile.name) $phoneFileSize != $($cloudFile.name) $cloudFileSize"
            $result = $false
        }
    }

    return $result
}

function CacheCloudFiles
{
    param([string]$cloudFolderPath, [string]$filter)

    $cloudFiles = @{}
    foreach ($file in Get-ChildItem -Path $cloudFolderPath -Recurse | where { $_.Name -match $filter }) {
        $cloudFiles.add($file.Name, $file)
    }

    return $cloudFiles
}

function GetConfig
{
    param([string]$configFile)

    if (![System.IO.File]::Exists($configFile)) {
        $configFile = [IO.Path]::Combine($PSScriptRoot, $configFile)
    }

    . $configFile

    return $config
}


$config = GetConfig($configFile)

if ($cloudFolderPath -eq "<configured>") {
    $cloudFolderPath = $config.settings.cloudFolderPath
}

if ($destinationFolderPath -eq "<configured>") {
    $destinationFolderPath = $config.settings.destinationFolderPath
}

$phoneInfo = $config.phones[$phoneName]
if ($phoneInfo -ne $null) {
    $phoneName = $phoneInfo.name
}

if ($phoneFolderPath -eq "<configured>") {
    if ($phoneInfo -eq $null) {
        throw "Failed to find $phoneName in $configFile"
    }
    $phoneFolderPath = $phoneInfo.folder
}

$phone = Get-Phone -phoneName $phoneName
if ($phone -eq $null) {
     throw "Can't find '$phoneName'. Have you attached the phone? Is it in 'File transfer' mode?"
}
if (!$(Test-Path -Path  $cloudFolderPath -PathType Container)) {
    throw "Can't find the folder '$cloudFolderPath'."
}
if (!$(Test-Path -Path  $destinationFolderPath -PathType Container)) {
    throw "Can't find the folder '$destinationFolderPath'."
}

# TODO add multiprocessing Write-Output "Running on $jobCount threads..."

$phoneFolder = Get-SubFolder -parent $phone -path $phoneFolderPath
Write-Output "Looking for files under '$phoneFolderPath' that match '$filter'..."
$phoneFiles = @( $phoneFolder.GetFolder.items() | where { $_.Name -match $filter } )
$phoneFileCount = $phoneFiles.Count
Write-Output "Found $phoneFileCount file(s) on the phone"

$phonePath = "$phoneName\$phoneFolderPath"
Write-Output "Processing path: $phonePath"
Write-Output "Looking for files in: $cloudFolderPath"
if ($phoneFileCount -gt 0) {
    $action = if ($dryRun) {"NOT copy (dry-run)"} else {"copy"}
    Write-Output "Will $action missing files to: $destinationFolderPath"

    $shell = Get-ShellProxy
    $destinationFolder = $shell.Namespace($destinationFolderPath).self

    Write-Output "Looking for files under '$cloudFolderPath' that match '$filter'..."
    $cloudFiles = CacheCloudFiles $cloudFolderPath $filter
    Write-Output "Found $($cloudFiles.Count) file(s) in the cloud"

    $processedCount = 0;
    $copied = 0;
    foreach ($phoneFile in $phoneFiles) {
        $fileName = $phoneFile.Name

        ++$processedCount
        $percent = [int](($processedCount * 100) / $phoneFileCount)
        Write-Progress -Activity "Processing Files in $phonePath" `
            -Status "Processing File ${count} / ${totalItems} (${percent}%)" `
            -CurrentOperation $fileName `
            -PercentComplete $percent

        if (!(IsInCloud $phoneFile $cloudFiles)) {
            if ($dryRun) {
                Write-Output "Would try to copy $fileName to $destinationFolderPath"
                ++$copied
            } else {
                $confirmed = $true
                if ($confirmCopy) {
                    $confirmation = Read-Host "$fileName seems to be missing from $cloudFolderPath"`
                                              " or any of its sub-folders."`
                                              " Shall we copy it to $destinationFolderPath? (y/n)"
                    if ($confirmation -ne 'y') {
                        $confirmed = $false
                    }
                }
                if ($confirmed) {
                    Write-Output "Copying $fileName to $destinationFolderPath..."
                    $destinationFolder.GetFolder.CopyHere($phoneFile)
                    ++$copied
                }
            }
        }
    }

    if ($copied -eq 0) {
        Write-Output "All $phoneFileCount file(s) seem to be already synced. 🎉🎉🎉"
    } else {
        $action = if ($dryRun) {"would have been"} else {"were"}
        Write-Output "$copied/$phoneFileCount item(s) $action copied to $destinationFolderPath"
    }
} else {
    Write-Output "Found no files under '$phonePath' matching the '$filter' filter."
}
