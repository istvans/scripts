<#
.SYNOPSIS
  Ensure all the files are uploaded from the phone to the cloud.
  Attempt to copy any missing or different files to a destination folder which
  is expected to be a cloud sync folder.
  If $beyondCompare can be resolved to a command, photos are compared with this
  tool before any copying might happen.
#>
param(
    [Parameter(Mandatory)]
    [string]$phoneName,
    [string]$phoneFolderPath="<configured>",
    [string]$cloudFolderPath="<configured>",
    [string]$destinationFolderPath="<configured>",
    [string]$filter=".(jpg|jpeg|mp4)$",
    [string]$beyondCompare="BComp.com",
    # See cloud_sync_config.example.ps1
    [string]$configFile="cloud_sync_config.ps1",
    [uint32]$jobCount=$(Get-ComputerInfo -Property CsProcessors).CsProcessors.NumberOfCores,
    [switch]$confirmCopy=$false,
    [switch]$dryRun=$false
)


#==============================================================================#
# JOBS
#==============================================================================#

$copyIfMissingJob = {
    param(
        [string]$phonePath,
        $phoneFile,
        [string]$cloudFolderPath,
        [string]$destinationFolderPath,
        [string]$beyondCompare,
        [bool]$confirmCopy,
        [bool]$dryRun
    )

    function GetMtpPath
    {
        param([string]$phonePath, [string]$fileName)

        $unixPath = $phonePath.Replace('\', '/')
        return "mtp://$unixPath/$fileName"
    }

    function FilesAreIdentical
    {
        param([string]$beyondCompare, [string]$phonePath, $phoneFile, [string]$cloudFile)

        $phoneFileMtpPath = GetMtpPath $phonePath $phoneFile.Name
        $arguments = "/silent /quickcompare `"$phoneFileMtpPath`" `"$cloudFile`""
        $process = Start-Process $beyondCompare -windowstyle Hidden -ArgumentList $arguments -PassThru -Wait
        $comparisonResult = $process.ExitCode

        $BINARY_SAME = 1
        $RULES_BASED_SAME = 2
        $SIMILAR = 12

        $result = switch ($comparisonResult) {
            {$_ -in $BINARY_SAME, $RULES_BASED_SAME, $SIMILAR} {
                $true
            }
            default {
                Write-Warning "Mismatch? => $beyondCompare $arguments return code: $comparisonResult"
                $false
            }
        }

        return $result
    }

    function GetExtendedProperty
    {
        param($phoneFile, [string]$property)

        try {
            return $phoneFile.ExtendedProperty($property)
        } catch [InvalidOperationException] {
            Write-Warning "$($phoneFile.Name) '$property' does not seem to be accessible"
            throw
        }
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
        $megaName = GetExtendedProperty($phoneFile, "System.DateModified").ToString("yyyy-MM-dd HH.mm.ss")
        $megaFilename = "$megaName$extension"
        return $megaFilename
    }

    function FindCloudFile
    {
        param([string]$cloudFolderPath, [string]$fileName)
        return @(Get-ChildItem -Path $cloudFolderPath -Recurse -Filter $originalFilename)[0]
    }

    function IsInCloud
    {
        param(
            [string]$phonePath,
            $phoneFile,
            [string]$cloudFolderPath,
            [string]$beyondCompare
        )

        $originalFilename = $phoneFile.name

        $cloudFile = FindCloudFile $cloudFolderPath $originalFilename

        if ($cloudFile -eq $null) {
            $megaFilename = GetMegaFilename($phoneFile)
            Write-Debug "Mega filename guess for '$originalFilename': '$megaFilename'"
            $cloudFile = FindCloudFile $cloudFolderPath $megaFilename
            Write-Debug "Mega filename match: $($cloudFile -ne $null)"
        }

        if ($cloudFile -eq $null) {
            $result = $false
        } else {
            $extension = [System.IO.Path]::GetExtension($originalFilename)
            Write-Debug "'$originalFilename' extension: '$extension'"

            if ($extension -eq ".mp4") {
                # These files have the exact same size both on the phone and on the PC.
                $phoneFileSize = GetExtendedProperty($phoneFile, "System.Size")
                $cloudFileSize = $cloudFile.Length
                Write-Debug "$($phoneFile.name) $phoneFileSize == $($cloudFile.name) $cloudFileSize"
                if ($phoneFileSize -eq $cloudFileSize) {
                    $result = $true
                } else {
                    $result = $false
                }
            } else {
                # These files can have different size or modified date on the phone and on the PC.
                # So we aim to do a proper comparison if we have a tool to do that.

                $command = Get-Command -Name $beyondCompare -ErrorAction SilentlyContinue
                if ($command -eq $null) {
                    $phoneFileSize = GetExtendedProperty $phoneFile "System.Size"
                    $cloudFileSize = $cloudFile.Length
                    if ($phoneFileSize -eq $cloudFileSize) {
                        $result = $true
                    } else {
                        $warningMsg = "'$cloudFile' seems to match '$originalFilename' but they have different sizes." + `
                                      " Without '$beyondCompare' we just assume that they are the same."
                        Write-Warning $warningMsg
                        $result = $true
                    }
                } else {
                    $result = FilesAreIdentical $beyondCompare $phonePath $phoneFile $cloudFile
                }
            }
        }

        return $result
    }

    # TODO DRY
    function GetShellProxy
    {
        if (-not $global:ShellProxy)
        {
            $global:ShellProxy = New-Object -Com Shell.Application
        }
        $global:ShellProxy
    }

    function CopyIfMissing
    {
        param(
            [string]$phonePath,
            $phoneFile,
            [string]$cloudFolderPath,
            [string]$destinationFolderPath,
            [string]$beyondCompare,
            [bool]$confirmCopy,
            [bool]$dryRun
        )

        if (IsInCloud $phonePath $phoneFile $cloudFolderPath $beyondCompare) {
            $copied = $false
        } else {
            $fileName = $phoneFile.Name

            if ($dryRun) {
                Write-Host "Would try to copy $fileName to $destinationFolderPath"
                $copied = $true
            } else {
                $confirmed = $true

                if ($confirmCopy) {
                    $confirmation = Read-Host "$fileName seems to be missing from $cloudFolderPath" `
                                              " or any of its sub-folders." `
                                              " Shall we copy it to $destinationFolderPath? (y/n)"
                    if ($confirmation -ne 'y') {
                        $confirmed = $false
                    }
                }

                if ($confirmed) {
                    Write-Host "Copying $fileName to $destinationFolderPath..."
                    $shell = GetShellProxy
                    $destinationFolder = $shell.Namespace($destinationFolderPath).self
                    $destinationFolder.GetFolder.CopyHere($phoneFile)
                    $copied = $true
                } else {
                    $copied = $false
                }
            }
        }

        return $copied
    }

    CopyIfMissing $phonePath $phoneFile $cloudFolderPath $destinationFolderPath $beyondCompare $confirmCopy $dryRun
}


#==============================================================================#
# FUNCTIONS
#==============================================================================#

function GetShellProxy
{
    if (-not $global:ShellProxy)
    {
        $global:ShellProxy = New-Object -Com Shell.Application
    }
    $global:ShellProxy
}

function GetPhone
{
    param($phoneName)
    $shell = GetShellProxy
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
    # => "My Computer" â€” the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
    # This folder can also contain mapped network drives.
    $shellItem = $shell.NameSpace(17).self
    $phone = $shellItem.GetFolder.items() | where { $_.name -eq $phoneName }
    return $phone
}

function GetPhoneSubFolder
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

function GetConfig
{
    param([string]$configFile)

    if (![System.IO.File]::Exists($configFile)) {
        $configFile = [IO.Path]::Combine($PSScriptRoot, $configFile)
    }

    . $configFile

    return $config
}

function WaitForJobsToFinish
{
    param($jobsStarted, $jobCount, $filesLeft, $processedCount, $phoneFileCount, $filePercent)

    if (($jobsStarted -eq $jobCount) -or ($filesLeft -eq $jobsStarted)) {
        while (Get-Job -State "Running") {
            $match = Get-Job | Select State | Select-String "Running"
            $running = $match.count
            $stillRunningJobsRatio = $running / $jobsStarted
            $jobPercent = [int]((1.0 - $stillRunningJobsRatio) * 100)

            Write-Progress -Activity "Waiting for $running/$jobsStarted job(s) to finish..." `
                           -Status "Processing files $processedCount / $phoneFileCount (${filePercent}%)" `
                           -PercentComplete $jobPercent
            Start-Sleep -Milliseconds 100
        }

        $copied = 0
        foreach ($job in Get-Job) {
            $wasCopied = Receive-Job -Job $job
            if ($wasCopied) {
                ++$copied
            }
        }

        Remove-Job *
        $jobsStarted = 0
    } else {
        $copied = 0
    }

    return $copied, $jobsStarted
}

#==============================================================================#
# MAIN
#==============================================================================#

if ($jobCount -eq 0) {
    throw "jobCount must be at least 1"
} elseif ($jobCount -gt 1 -and $confirmCopy) {
    throw "jobCount must be 1 if confirmCopy is enabled"
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

$phone = GetPhone $phoneName
if ($phone -eq $null) {
     throw "Can't find '$phoneName'. Have you attached the phone? Is it in 'File transfer' mode?"
}
if (!$(Test-Path -Path  $cloudFolderPath -PathType Container)) {
    throw "Can't find the folder '$cloudFolderPath'."
}
if (!$(Test-Path -Path  $destinationFolderPath -PathType Container)) {
    throw "Can't find the folder '$destinationFolderPath'."
}

$phoneFolder = GetPhoneSubFolder $phone $phoneFolderPath
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

    Write-Output "Remove any leftover job..."
    Remove-Job -Force *

    if ($jobCount -eq 1) {
        Write-Output "Running on a single thread..."
    } else {
        Write-Output "Running on $jobCount threads..."
    }

    $processedCount = 0
    $copied = 0
    $jobsStarted = 0
    $filesLeft = $phoneFileCount
    foreach ($phoneFile in $phoneFiles) {
        $fileName = $phoneFile.Name

        ++$processedCount
        $percent = [int](($processedCount * 100) / $phoneFileCount)
        Write-Progress -Activity "Processing files under $phonePath" `
            -Status "Processing file $processedCount / $phoneFileCount (${percent}% copied:$copied)" `
            -CurrentOperation $fileName `
            -PercentComplete $percent

        --$filesLeft

        $arguments = $phonePath, $phoneFile, $cloudFolderPath, $destinationFolderPath, $beyondCompare, $confirmCopy, $dryRun
        if ($jobCount -eq 1) {
            $wasCopied = Invoke-Command -ScriptBlock $copyIfMissingJob -ArgumentList $arguments
            if ($wasCopied) {
                ++$copied
            }
        } else {
            Start-Job -ScriptBlock $copyIfMissingJob -ArgumentList $arguments >$null
            ++$jobsStarted

            # $beyondCompare doesn't seem to like to start multiple instances
            # all at once?
            Start-Sleep -MilliSeconds 100

            $newlyCopied, $jobsStarted = WaitForJobsToFinish $jobsStarted $jobCount $filesLeft `
                $processedCount $phoneFileCount $percent
            $copied += $newlyCopied
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
