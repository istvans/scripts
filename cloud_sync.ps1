<#
.SYNOPSIS
  Ensure all the files are uploaded from the phone to the cloud.
  Attempt to copy any missing or different files to a destination folder which
  is expected to be a cloud sync folder.
#>
param(
    [Parameter(Mandatory)]
    [string]$phoneName,
    [string]$phoneFolderPath="<configured>",
    [string]$cloudFolderPath="<configured>",
    [string]$destinationFolderPath="<configured>",
    [string]$filter=".(jpg|jpeg|mp4)$",
    [string]$choco="choco",
    [string]$winmerge="C:\Program Files\WinMerge\WinMergeU.exe",
    [string]$tempDir=$env:Temp,
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
        $phoneFile,
        [string]$cloudFolderPath,
        [string]$tempDir,
        [string]$destinationFolderPath,
        [string]$winmerge,
        [switch]$confirmCopy,
        [switch]$dryRun
    )

    function GetTempFilePath
    {
        param($phoneFile, [string]$tempDir)

        return [IO.Path]::Combine($tempDir, $phoneFile.Name)
    }

    # TODO do not copy
    function GetShellProxy
    {
        if (-not $global:ShellProxy)
        {
            $global:ShellProxy = New-Object -Com Shell.Application
        }
        return $global:ShellProxy
    }

    function CreateTempFile
    {
        param($phoneFile, [string]$tempDir)

        $shell = GetShellProxy
        $destinationFolder = $shell.Namespace($tempDir).self

        $destinationFolder.GetFolder.CopyHere($phoneFile)

        $tempFilePath = GetTempFilePath $phoneFile $tempDir
        if ([System.IO.File]::Exists($tempFilePath)) {
            Write-Debug "Created temporary file '$tempFilePath'"
        } else {
            throw "Failed to create '$tempFilePath'"
        }

        return $tempFilePath
    }

    function FilesAreIdentical
    {
        param([string]$winmerge, $phoneFile, $cloudFile)

        $arguments = "-enableexitcode -noninteractive -minimize `"$phoneFile`" `"$cloudFile`""
        $process = Start-Process $winmerge -windowstyle Hidden -ArgumentList $arguments -PassThru -Wait
        $comparisonResult = $process.ExitCode

        enum Result {
            Identical = 0
            Different
            Error
        }
        $enumResult = [Enum]::ToObject([Result], $comparisonResult)

        $result = switch ($enumResult) {
            Identical { $true }
            Different { $false }
            Error {
                throw "Failed to compare '$phoneFile' and '$cloudFile'"
            }
        }

        return $result
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

    function FindCloudFile
    {
        param([string]$cloudFolderPath, [string]$fileName)
        return @(Get-ChildItem -Path $cloudFolderPath -Recurse -Filter $originalFilename)[0]
    }

    function IsInCloud
    {
        param(
            $phoneFile,
            [string]$cloudFolderPath,
            [string]$tempDir,
            [string]$winmerge
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
                # These files have the exact same size on the phone and on the PC too.
                $phoneFileSize = $phoneFile.ExtendedProperty("System.Size")
                $cloudFileSize = $cloudFile.Length
                Write-Debug "$($phoneFile.name) $phoneFileSize == $($cloudFile.name) $cloudFileSize"
                if ($phoneFileSize -eq $cloudFileSize) {
                    $result = $true
                } else {
                    $result = $false
                }
            } else {
                # These files can have different size or modified date on the phone and on the PC.
                # So we do a proper comparison using winmerge.

                $tempFile = CreateTempFile $phoneFile $tempDir
                $result = FilesAreIdentical $winmerge $tempFile $cloudFile
            }
        }

        return $result
    }

    function CopyIfMissing
    {
        param(
            $phoneFile,
            [string]$cloudFolderPath,
            [string]$tempDir,
            [string]$destinationFolderPath,
            [string]$winmerge,
            [switch]$confirmCopy,
            [switch]$dryRun
        )

        if (IsInCloud $phoneFile $cloudFolderPath $tempDir $winmerge) {
            $copied = $false
        } else {
            $fileName = $phoneFile.Name

            if ($dryRun) {
                Write-Output "Would try to copy $fileName to $destinationFolderPath"
                $copied = $true
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
                    # re-use the temporary file if it was created for the comparison
                    $tempFilePath = GetTempFilePath $phoneFile $tempDir
                    if ([System.IO.File]::Exists($tempFilePath)) {
                        $shell = GetShellProxy
                        $tempFile = $shell.Namespace($tempFilePath).self
                        $fileToCopy = $tempFile
                    } else {
                        $fileToCopy = $phoneFile
                    }

                    Write-Output "Copying $fileName to $destinationFolderPath..."
                    $shell = GetShellProxy
                    $destinationFolder = $shell.Namespace($destinationFolderPath).self
                    $destinationFolder.GetFolder.CopyHere($fileToCopy)
                    $copied = $true
                } else {
                    $copied = $false
                }
            }
        }

        return $copied
    }

    CopyIfMissing $phoneFile $cloudFolderPath $tempDir $destinationFolderPath $winmerge $confirmCopy $dryRun
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

function EnsureChocoIsInstalled
{
    param([string]$choco)

    $command = Get-Command -Name $choco -ErrorAction SilentlyContinue
    if ($command -eq $null) {
        Write-Output "'$choco' is missing. Let's install it."

        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        $chocoInstaller = "https://community.chocolatey.org/install.ps1"
        $arguments = "Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('$chocoInstaller'))"
        Start-Process powershell -ArgumentList $arguments -Verb RunAs -Wait
    } else {
        Write-Output "'$choco' is installed."
    }
}

function EnsureWinMergeIsInstalled
{
    param([string]$choco, [string]$winmerge)

    $command = Get-Command -Name $winmerge -ErrorAction SilentlyContinue
    if ($command -eq $null) {
        Write-Output "'$winmerge' is missing. Let's install it."

        EnsureChocoIsInstalled $choco

        # handle the case when choco was freshly installed
        $command = Get-Command -Name $choco -ErrorAction SilentlyContinue
        if ($command -eq $null) {
            $choco = [IO.Path]::Combine($env:ChocolateyInstall, "choco")
            $command = Get-Command -Name $choco -ErrorAction SilentlyContinue
            if ($command -eq $null) {
                throw "Cannot find '$choco'"
            }
        }

        Start-Process $choco -ArgumentList "install winmerge" –Verb RunAs -Wait
    } else {
        Write-Output "'$winmerge' is installed."
    }
}

function EnsurePrerequisites
{
    param([string]$choco, [string]$winmerge)

    EnsureWinMergeIsInstalled $choco $winmerge
}

function ClearTempDirectory
{
    param([string]$tempDir, $filter)

    Write-Output "Cleaning '$tempDir'..."
    Get-ChildItem -Path $tempDir -File | where { $_.Name -match $filter } | foreach { $_.Delete()}
    Write-Output "Done"
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
    } else {
        $copied = 0
    }

    return $copied
}

#==============================================================================#
# MAIN
#==============================================================================#

if ($jobCount -eq 0) {
    throw "jobCount must be at least 1"
}

EnsurePrerequisites $choco $winmerge

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

    # Write-Output "Looking for files under '$cloudFolderPath' that match '$filter'..."
    # $cloudFiles = CacheCloudFiles $cloudFolderPath $filter
    # Write-Output "Found $($cloudFiles.Count) file(s) in the cloud"

    ClearTempDirectory $tempDir $filter  # this helps re-runs

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

        $arguments = $phoneFile, $cloudFolderPath, $tempDir, $destinationFolderPath, $winmerge, $confirmCopy, $dryRun
        if ($jobCount -eq 1) {
            $wasCopied = Invoke-Command -ScriptBlock $copyIfMissingJob -ArgumentList $arguments
            if ($wasCopied) {
                ++$copied
            }
        } else {
            Start-Job $copyIfMissingJob -ArgumentList $arguments >$null
            ++$jobsStarted

            $copied += WaitForJobsToFinish $jobsStarted $jobCount $filesLeft $processedCount $phoneFileCount $percent
        }

        if ($processedCount -eq 4) {
            Write-Debug "DEBUG EXIT NOW"
            exit 42
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
