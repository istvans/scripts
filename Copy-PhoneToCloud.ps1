<#
.SYNOPSIS
  Ensure the phone files are uploaded to "the cloud".

.DESCRIPTION
  Find all the files on the `PhoneName` phone under `PhoneFolderPath` applying
  the file `Filter`. Try to find all these files, if any, under the
  `CloudFolderPath` directory. Attempt to copy any missing or different files to
  the `DestinationFolderPath` directory which is expected to be a cloud sync
  folder (hence the name).

  If `BeyondCompare` can be resolved to a command, files are compared with this
  tool before any copying might happen. If content difference is found, the copy
  will be attempted.
  Without `BeyondCompare` the script only compares the file sizes and assumes
  a difference if the file sizes do not match. This will likely trigger
  unnecessary copies. You can use the `OnlyCompareFileNames` switch to disable
  this.

  If the script finds content/size difference between files with matching names
  the attempted copy will try to overwrite the file in the `DestinationFolderPath`
  iff that file exists there (e.g. `DestinationFolderPath == `CloudFolderPath``).
  In this case Windows will request a confirmation and should provide the means
  to compare the two files.

  Suggestion: If `CloudFolderPath` and `DestinationFolderPath` are different
  folders and `DestinationFolderPath` is an empty folder, you can easily implement
  a two stage copy, where you'd have all the files collected into
  `DestinationFolderPath`. This can be handy if you expect a lot of conflict to
  avoid dealing with a lot of pop-up windows during the script's run (could be
  especially annoying with a lot of threads).

  Mega.io cloud support
  ---------------------
  If a phone file isn't available under `CloudFolderPath`, the script tries a
  very basic, automatic name-conversion following the format mega.io might use
  if "Keep file names as in the device" was disabled when the file was uploaded.
  This method cannot deal with things like time zone changes.

.PARAMETER PhoneName
  The name of the attached phone. This can be the "nickname" of the phone to
  identify it in the `ConfigFile` (see below) or the actual name of the phone as
  it shows up in the File Explorer e.g. "ONEPLUS A6013" (without quotes).

.PARAMETER PhoneFolderPath
  The path to a folder on the `PhoneName` phone where the script should look for
  files.

  Can be `"<configured>"`, see `ConfigFile`.

.PARAMETER CloudFolderPath
  The script tries to find the found phone files under this folder. If a file can
  be found here, the script won't copy it.

  Can be `"<configured>"`, see `ConfigFile`.

.PARAMETER DestinationFolderPath
  If a file couldn't be found under `CloudFolderPath` or its content is
  sufficiently different from the file under `CloudFolderPath` and `DryRun` was
  not set, then the script will attempt to copy the file into this directory.

  Note: The script needs to use a COM object for the copy and the `CopyHere`
  method does not return if the copy occurred or not. So the script will log
  every attempted copy as a copy no matter what.

  Can be `"<configured>"`, see `ConfigFile`.

.PARAMETER Filter
  The phone file filter. It's a regular expression that needs to match each file
  that we would like to make sure they are uploaded to the cloud.

.PARAMETER BeyondCompare
  The name of `BComp.com` if the application is on the `PATH` or the full path
  to the application.

  This tool is used to determine if two files with the same name also have the
  same content.

  Note: The script does not use this tool for video files (this just means `.mp4`
  at the moment). For such files comparing the size seems adequate i.e. the file
  size do not seem to vary between the Phone and Windows (phone filesystem vs NTFS).

  You need a fully licenced version of the tool. Testing with a trial version
  suggested the tool would fail asking for a licence key in this scenario i.e.
  potentially running from multiple threads. The single threaded operation might
  work even with a trial version of `BeyondCompare` but this wasn't explicitly
  tested.

.PARAMETER ConfigFile
  Make the use of the script simpler by allowing the User to just specify the
  "nickname" of their phone and leave every other parameter on their default
  values. The parameters with "<configured>" default value will try to read their
  actual value from the ConfigFile.

  The ConfigFile should be a powershell script with the following structure:
  ```powershell
    $config = @{
        settings = @{
            cloudFolderPath = "<your local cloud sync folder>"
            destinationFolderPath = "<where to copy missing files to>"
        }
        phones = @{
            oneplus = @{
                name = "<your phone's name as it is shown in the File Explorer>"
                folder = "<the full path (without the phone name) on your phone to"`
                         " the folder where you want to sync files from>"
            }
            samsung = @{
                # ...
            }
            iphone = @{
                # ...
            }
        }
    }
  ```

  If the ConfigFile does not exist and any of the arguments have the `"<configured>"`
  value, the script will fail. Otherwise, the script will try to use `PhoneName`
  as the real phone name and all the other arguments as ones with proper values.

  In other words: the script can be used with or without a `ConfigFile`.

.PARAMETER ThreadCount
  By default this is set to the number of cores of the CPU in the machine that
  is executing the script. Its value must be at least 1.

  The script displays a progress bar to show the total progress and a progress
  bar for each thread to visualise their individual progresses. This latter
  feature is limited to 45 threads, to keep the progress bars displayable even
  on smaller screens.

.PARAMETER OnlyCompareFileNames
  Do not compare the content or size of any file. Can be handy if you don't have
  `BeyondCompare`. A file is only copied to the `DestinationFolderPath` if there
  is no file matching its name in the `CloudFolderPath`.

  Note: The mega.io name check is still enabled in this case.

  Apart from that and the multithreading support, this mode is not much different
  from trying to copy all your files from `PhoneFolderPath` to `CloudFolderPath`
  if `CloudFolderPath` equals `DestinationFolderPath` using File Explorer.

.PARAMETER DryRun
  Perform all the steps except the actual copying. Extremely handy for debugging.

#>
param(
    [Parameter(Mandatory)]
    [string]$PhoneName,
    [string]$PhoneFolderPath = "<configured>",
    [string]$CloudFolderPath = "<configured>",
    [string]$DestinationFolderPath = "<configured>",
    [string]$Filter = ".(jpg|jpeg|mp4)$",
    [string]$BeyondCompare = "BComp.com",
    [string]$ConfigFile = "copy_phone_to_cloud_config.ps1",
    [uint32]$ThreadCount = $(Get-ComputerInfo -Property CsProcessors).CsProcessors.NumberOfCores,
    [switch]$OnlyCompareFileNames = $false,
    [switch]$DryRun = $false
)


#==============================================================================#
# Functions
#==============================================================================#

function Get-Phone {
    param($PhoneName)

    $shell = New-Object -ComObject Shell.Application
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
    # => "My Computer" â€” the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
    # This folder can also contain mapped network drives.
    $shellItem = $shell.NameSpace(17).self
    $phone = $shellItem.GetFolder.items() | Where-Object { $_.name -eq $PhoneName }
    return $phone
}


function Get-PhoneSubFolder {
    param($parent, [string]$path)
    $pathParts = @( $path.Split([system.io.path]::DirectorySeparatorChar) )
    $current = $parent
    foreach ($pathPart in $pathParts) {
        if ($pathPart) {
            $current = $current.GetFolder.items() | Where-Object { $_.Name -eq $pathPart }
        }
    }
    return $current
}


function Get-Config {
    param([string]$ConfigFile)

    if (![System.IO.File]::Exists($ConfigFile)) {
        $ConfigFile = [IO.Path]::Combine($PSScriptRoot, $ConfigFile)
    }

    if ([System.IO.File]::Exists($ConfigFile)) {
        . $ConfigFile
    }
    else {
        $config = $null
    }

    return $config
}


function Get-TotalProcessedAndCopied {
    param($StatusArray)

    $totalProcessed = $StatusArray | ForEach-Object { $_.PhoneFileProcessed } | Measure-Object -Sum | ForEach-Object Sum
    $totalCopied = $StatusArray | ForEach-Object { $_.PhoneFileCopied } | Measure-Object -Sum | ForEach-Object Sum
    $totalProcessed, $totalCopied
}


function Write-AllProgress {
    param($StatusArray, [Int]$Goal)

    $totalProcessed, $totalCopied = Get-TotalProcessedAndCopied $StatusArray

    $parentProgressId = 0
    $percent = [int]($totalProcessed * 100 / $Goal)
    Write-Progress -Id $parentProgressId -Activity "Total:" `
        -Status "Copied: $totalCopied Processed: $totalProcessed/$Goal" `
        -PercentComplete $percent

    # Arbitrary limit to prevent too many progress bars that can't be properly
    # displayed, at least on relatively small screens.
    $limit = 45
    if ($StatusArray.Length -le $limit) {
        for ($threadId = 0; $threadId -lt $StatusArray.Length; ++$threadId) {
            $runspaceStatus = $StatusArray[$threadId]
            $processed = $runspaceStatus.PhoneFileProcessed
            $copied = $runspaceStatus.PhoneFileCopied
            $threadProgressId = $threadId + 1  # the ParentId can't be the same as the Id
            $threadPercent = [int]($processed * 100 / $Goal)  # will never reach 100, but that's fine
            Write-Progress -ParentId $parentProgressId -Id $threadProgressId `
                -Activity "Thread${threadProgressId}:" `
                -Status "Copied: $copied Processed: $processed/$Goal" `
                -PercentComplete $threadPercent
        }
    }
}


#==============================================================================#
# Threading
#==============================================================================#

function Invoke-ThreadTop {
    param (
        [string]$PhonePath,
        [string]$CloudFolderPath,
        [string]$DestinationFolderPath,
        [string]$BeyondCompare,
        [bool]$DryRun,
        [bool]$OnlyCompareFileNames,
        [System.Collections.Concurrent.ConcurrentQueue[System.__ComObject]]$InputQueue,
        [System.Collections.Concurrent.ConcurrentQueue[String]]$OutputQueue,
        [Array]$StatusArray,
        [Int]$ThreadId,
        [bool]$KeepOnRunning
    )

    # Return a string that BeyondCompare can consume as an MTP path to the file.
    function Get-MtpPath {
        param([string]$PhonePath, [string]$FileName)

        $unixPath = $PhonePath.Replace('\', '/')
        return "mtp://$unixPath/$FileName"
    }


    function Test-FilesAreIdentical {
        param(
            [string]$BeyondCompare,
            [string]$PhonePath,
            [System.__ComObject]$PhoneFile,
            [string]$CloudFile
        )

        $phoneFileMtpPath = Get-MtpPath $PhonePath $PhoneFile.Name
        $arguments = "/silent /quickcompare `"$phoneFileMtpPath`" `"$CloudFile`""
        $process = Start-Process $BeyondCompare -WindowStyle Hidden -ArgumentList $arguments -PassThru -Wait
        $comparisonResult = $process.ExitCode

        $BINARY_SAME = 1
        $RULES_BASED_SAME = 2
        $SIMILAR = 12
        $ERROR_CODE = 100

        $result = switch ($comparisonResult) {
            { $_ -in $BINARY_SAME, $RULES_BASED_SAME, $SIMILAR } {
                $true
            }
            { $_ -eq $ERROR_CODE } {
                Write-Error "$BeyondCompare $arguments return code: $comparisonResult"
                $true
            }
            default {
                Write-Warning "Mismatch? => $BeyondCompare $arguments return code: $comparisonResult"
                $false
            }
        }

        return $result
    }


    function Get-ExtendedProperty {
        param([System.__ComObject]$phoneFile, [string]$property)

        try {
            return $phoneFile.ExtendedProperty($property)
        }
        catch [InvalidOperationException] {
            Write-Warning "$($phoneFile.Name) '$property' does not seem to be accessible"
            throw
        }
    }


    # Return a possible equivalent of the name of `phoneFile` in the cloud when
    # "Keep file names as in the device" is disabled.
    function Get-MegaFilename {
        param([System.__ComObject]$PhoneFile)

        $extension = [System.IO.Path]::GetExtension($PhoneFile.Name)
        $dateModified = Get-ExtendedProperty $PhoneFile "System.DateModified"
        $megaName = $dateModified.ToString("yyyy-MM-dd HH.mm.ss")
        $megaFilename = "$megaName$extension"
        return $megaFilename
    }


    function Find-CloudFile {
        param([string]$CloudFolderPath, [string]$FileName)
        return @(Get-ChildItem -Path $CloudFolderPath -Recurse -Filter $FileName)[0]
    }


    function Test-FileIsInCloud {
        param(
            [string]$PhonePath,
            [System.__ComObject]$PhoneFile,
            [string]$CloudFolderPath,
            [string]$BeyondCompare,
            [bool]$OnlyCompareFileNames
        )

        $originalFilename = $PhoneFile.Name

        $cloudFile = Find-CloudFile -CloudFolderPath $CloudFolderPath -FileName $originalFilename

        if ($cloudFile -eq $null) {
            $megaFilename = Get-MegaFilename $PhoneFile
            Write-Debug "Mega filename guess for '$originalFilename': '$megaFilename'"
            $cloudFile = FindCloudFile $CloudFolderPath $megaFilename
            Write-Debug "Mega filename match: $($cloudFile -ne $null)"
        }

        if ($cloudFile -eq $null) {
            $result = $false
        }
        elseif ($OnlyCompareFileNames) {
            $result = $true
        }
        else {
            $extension = [System.IO.Path]::GetExtension($originalFilename)
            Write-Debug "'$originalFilename' extension: '$extension'"

            if ($extension -eq ".mp4") {
                # These files have the exact same size both on the phone and on the PC.
                $phoneFileSize = Get-ExtendedProperty $PhoneFile "System.Size"
                $cloudFileSize = $cloudFile.Length
                Write-Debug "$($PhoneFile.name) $phoneFileSize == $($cloudFile.Name) $cloudFileSize"
                if ($phoneFileSize -eq $cloudFileSize) {
                    $result = $true
                }
                else {
                    $result = $false
                }
            }
            else {
                # These files can have different size or modified date on the phone
                # and on the PC.
                # So we aim to do a proper comparison if we have a tool to do that.

                $command = Get-Command -Name $BeyondCompare -ErrorAction SilentlyContinue
                if ($command -eq $null) {
                    $phoneFileSize = Get-ExtendedProperty $PhoneFile "System.Size"
                    $cloudFileSize = $cloudFile.Length
                    if ($phoneFileSize -eq $cloudFileSize) {
                        $result = $true
                    }
                    else {
                        $warningMsg = "'$cloudFile' and '$originalFilename' might be the same file but they" + `
                            " have different sizes: $phoneFileSize vs $cloudFileSize." + `
                            " Without '$BeyondCompare' we just assume that they are different."
                        Write-Warning $warningMsg
                        $result = $false
                    }
                }
                else {
                    $result = Test-FilesAreIdentical $BeyondCompare $PhonePath $PhoneFile $cloudFile
                }
            }
        }

        return $result
    }


    function Copy-IfMissing {
        param(
            [string]$PhonePath,
            [System.__ComObject]$PhoneFile,
            [string]$CloudFolderPath,
            [string]$DestinationFolderPath,
            [string]$BeyondCompare,
            [bool]$OnlyCompareFileNames,
            [bool]$DryRun,
            [System.__ComObject]$Shell
        )

        if (Test-FileIsInCloud $PhonePath $PhoneFile $CloudFolderPath $BeyondCompare $OnlyCompareFileNames) {
            $copied = $false
        }
        else {
            $fileName = $PhoneFile.Name

            if ($DryRun) {
                Write-Host "Would try to copy $fileName to $DestinationFolderPath"
                $copied = $true
            }
            else {
                Write-Host "Copying $fileName to $DestinationFolderPath..."
                $destinationFolder = $Shell.Namespace($DestinationFolderPath).self
                $destinationFolder.GetFolder.CopyHere($PhoneFile)
                $copied = $true
            }
        }

        return $copied
    }


    function Invoke-ProcessLoop {
        param (
            [string]$PhonePath,
            [string]$CloudFolderPath,
            [string]$DestinationFolderPath,
            [string]$BeyondCompare,
            [bool]$DryRun,
            [bool]$OnlyCompareFileNames,
            [System.Collections.Concurrent.ConcurrentQueue[System.__ComObject]]$InputQueue,
            [System.Collections.Concurrent.ConcurrentQueue[String]]$OutputQueue,
            [Array]$StatusArray,
            [Int]$ThreadId,
            [bool]$KeepOnRunning
        )

        $shell = New-Object -ComObject Shell.Application

        $processed = 0
        $copied = 0
        $phoneFile = $null
        while ($KeepOnRunning.Value) {
            if ($InputQueue.TryDequeue([ref]$phoneFile)) {
                if ($phoneFile -eq $null) {
                    $status = "Processed a NULL phoneFile on thread $ThreadId..."
                    ++$processed
                    $StatusArray[$ThreadId].PhoneFileProcessed = $processed
                }
                else {
                    $status = "Processing $($phoneFile.Name) on thread $ThreadId..."

                    $wasCopied = Copy-IfMissing -PhonePath $PhonePath -PhoneFile $phoneFile `
                        -CloudFolderPath $CloudFolderPath -DestinationFolderPath $DestinationFolderPath `
                        -BeyondCompare $BeyondCompare -OnlyCompareFileNames $OnlyCompareFileNames `
                        -DryRun $DryRun -Shell $shell
                    if ($wasCopied) {
                        ++$copied
                    }
                    ++$processed

                    $StatusArray[$ThreadId].PhoneFileProcessed = $processed
                    $StatusArray[$ThreadId].PhoneFileCopied = $copied
                }
            }
            else {
                $waitTimeMs = 100
                $status = "Thread $ThreadId is waiting $waitTimeMs for a new message..."
                Start-Sleep -Milliseconds $waitTimeMs
            }

            $OutputQueue.Enqueue($status)
        }
    }

    $arguments = @{
        PhonePath = $PhonePath
        CloudFolderPath = $CloudFolderPath
        DestinationFolderPath = $DestinationFolderPath
        BeyondCompare = $BeyondCompare
        DryRun = $DryRun
        OnlyCompareFileNames = $OnlyCompareFileNames
        InputQueue = $InputQueue
        OutputQueue = $OutputQueue
        StatusArray = $StatusArray
        ThreadId = $ThreadId
        KeepOnRunning = $KeepOnRunning
    }
    Invoke-ProcessLoop @arguments
}


#==============================================================================#
# Types
#==============================================================================#

class RunspaceStatus {
    [Int]$PhoneFileProcessed
    [Int]$PhoneFileCopied

    RunspaceStatus() {
        $this.PhoneFileProcessed = 0
        $this.PhoneFileCopied = 0
    }
}


#==============================================================================#
# Main
#==============================================================================#

if ($ThreadCount -eq 0) {
    throw "ThreadCount must be at least 1"
}

$config = Get-Config($ConfigFile)

if ($config -eq $null) {
    Write-Output "'$ConfigFile' is not a valid config."
}
else {
    # Substitute placeholders with the configured, real values.
    if ($CloudFolderPath -eq "<configured>") {
        $CloudFolderPath = $config.settings.cloudFolderPath
    }

    if ($DestinationFolderPath -eq "<configured>") {
        $DestinationFolderPath = $config.settings.destinationFolderPath
    }

    $phoneInfo = $config.phones[$PhoneName]
    if ($phoneInfo -ne $null) {
        $PhoneName = $phoneInfo.name
    }

    if ($PhoneFolderPath -eq "<configured>") {
        if ($phoneInfo -eq $null) {
            throw "Failed to find $PhoneName in $ConfigFile"
        }
        $PhoneFolderPath = $phoneInfo.folder
    }
}

$phone = Get-Phone $PhoneName
if ($phone -eq $null) {
    throw "Can't find '$PhoneName'. Have you attached the phone? Is it in 'File transfer' mode? Did you specify a valid config?"
}
if (!$(Test-Path -Path  $CloudFolderPath -PathType Container)) {
    throw "Can't find the folder '$CloudFolderPath'."
}
if (!$(Test-Path -Path  $DestinationFolderPath -PathType Container)) {
    throw "Can't find the folder '$DestinationFolderPath'."
}

$phoneFolder = Get-PhoneSubFolder $phone $PhoneFolderPath
Write-Output "Looking for files under '$PhoneFolderPath' that match '$Filter'..."
$phoneFiles = @( $phoneFolder.GetFolder.items() | Where-Object { $_.Name -match $Filter } )
$phoneFileCount = $phoneFiles.Count
Write-Output "Found $phoneFileCount file(s) on the phone"

$phonePath = "$PhoneName\$PhoneFolderPath"
Write-Output "Processing path: $phonePath"
Write-Output "Looking for files in: $CloudFolderPath"
if ($phoneFileCount -gt 0) {
    $action = if ($DryRun) { "NOT copy (dry-run)" } else { "copy" }
    Write-Output "Will $action missing files to: $DestinationFolderPath"

    if ($ThreadCount -eq 1) {
        Write-Output "Running on a single thread..."
    }
    else {
        Write-Output "Running on $ThreadCount threads..."
    }

    $runspaces = 1..$ThreadCount | Foreach-Object { [PowerShell]::Create() }
    $runspaceStatuses = 1..$ThreadCount | Foreach-Object { [RunspaceStatus]::new() }
    $inputQueue = New-Object -TypeName "System.Collections.Concurrent.ConcurrentQueue[System.__ComObject]"
    $debugOutputQueue = New-Object -TypeName "System.Collections.Concurrent.ConcurrentQueue[String]"

    # disable ctrl + c so we can handle it manually
    [Console]::TreatControlCAsInput = $true

    Write-Output "Start $ThreadCount thread(s)..."
    $keepOnRunning = $true
    for ($threadId = 0; $threadId -lt $ThreadCount; ++$threadId) {
        $null = $runspaces[$threadId].AddScript(${function:Invoke-ThreadTop}).
        AddParameter("PhonePath", $phonePath).
        AddParameter("CloudFolderPath", $CloudFolderPath).
        AddParameter("DestinationFolderPath", $DestinationFolderPath).
        AddParameter("BeyondCompare", $BeyondCompare).
        AddParameter("DryRun", $DryRun).
        AddParameter("OnlyCompareFileNames", $OnlyCompareFileNames).
        AddParameter("InputQueue", $inputQueue).
        AddParameter("OutputQueue", $debugOutputQueue).
        AddParameter("StatusArray", $runspaceStatuses).
        AddParameter("ThreadId", $threadId).
        AddParameter("KeepOnRunning", [ref]$keepOnRunning).BeginInvoke()
    }

    try {
        Write-Output "Queue $phoneFileCount phone file(s)..."
        foreach ($phoneFile in $phoneFiles) {
            $inputQueue.Enqueue($phoneFile)
        }
        Write-Output "Queue size: $($inputQueue.Count)"

        $threadingChoice = if ($ThreadCount -eq 1) { "Single" } else { "Multi" }
        Write-Output "${threadingChoice}-threaded processing..."
        do {
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.key -eq "C" -and $key.modifiers -eq "Control") {
                    Write-Output "Cancelled"
                    break
                }
            }

            Write-AllProgress -StatusArray $runspaceStatuses -Goal $phoneFileCount

            if ($Debug) {
                $scriptOutput = $null
                while ($debugOutputQueue.TryDequeue([ref]$scriptOutput)) {
                    $scriptOutput
                }
            }

            # If the threads are done processing the input we gave them, let them know they can stop
            if ($inputQueue.IsEmpty) {
                $keepOnRunning = $false
            }

            Start-Sleep -Milliseconds 100

            ## See if we still have any busy runspaces. If not, exit the loop.
            $busyRunspaces = $runspaces | Where-Object { $_.InvocationStateInfo.State -ne 'Complete' }
        } while ($busyRunspaces)

        # Make sure we've got the final numbers
        $totalProcessed, $totalCopied = Get-TotalProcessedAndCopied $runspaceStatuses

        # Final result and wrapping up.
        if ($totalProcessed -eq $phoneFileCount) {
            if ($totalCopied -eq 0) {
                Write-Output "All $phoneFileCount file(s) seem to be already synced. 🎉🎉🎉"
            }
            else {
                $action = if ($DryRun) { "would have been" } else { "were" }
                Write-Output "$totalCopied/$phoneFileCount item(s) $action copied to $DestinationFolderPath"
            }
        }
        else {
            Write-Output "Processed $totalProcessed from $phoneFileCount"
        }
    }
    finally {
        Write-Output "Stop the $($runspaces.Length) thread(s)..."
        foreach ($runspace in $runspaces) {
            $runspace.Stop()
            $runspace.Dispose()
        }
    }
}
else {
    Write-Output "Found no files under '$phonePath' matching the '$Filter' filter."
}
