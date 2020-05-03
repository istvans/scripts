# Sync all *.cloud files/directories in the specified `OdrivePath`.
# TODO:
# =====
# - Add multiple runs thus newly added multiple levels of folders can be synced
#   in one go.
param(
    [Parameter(ValueFromPipeline)][System.Array] $CloudFiles,
    [String] $OdrivePath = "D:\odrive",
    [String] $OdriveScript = "$HOME\.odrive-agent\bin\odrive.py",
    [String] $Python = "python",
    [String] $ErrorFile = "errors.txt",
    [String] $ExcludePattern = $null,
    [String] $IncludePattern = $null,
    [uint32] $JobCount = 8,
    [switch] $LiveErrorReporting,
    [switch] $JustFind
)

$global:multi_threaded = if ($JobCount -ne 0) {$True} else {$False}
$mt_status = if ($global:multi_threaded) {"enabled"} else {"disabled"}
Write-Host "Multi-threaded operation is $mt_status"

if (!(Test-Path "$OdrivePath" -PathType Container)) {
    throw "Unable to find '$OdrivePath'"
}

if (!(Test-Path "$OdriveScript" -PathType Leaf)) {
    throw "Unable to find '$OdriveScript'"
}

if ((Get-Command "$Python" -ErrorAction SilentlyContinue) -eq $null) {
    throw "Unable to find '$Python' in your PATH"
}

function find_cloud_files([String]$path)
{
    Write-Host "Gathering cloud files in '$path'... Please wait... " -NoNewLine
    return ls -Include *.cloud* -Path $path -Recurse
}

if ($CloudFiles.length -eq 0) {
    $CloudFiles = find_cloud_files $OdrivePath
}

$global:number_of_files = $CloudFiles.length
Write-Host "Found $global:number_of_files cloud files to sync"

if ($JustFind) {
    Write-Host "JustFind was specified, so just returning the found files"
    return $CloudFiles
}

function update_status([uint32]$counter,
                       [DateTime]$start_time,
                       [String]$percent_string,
                       [uint32]$done_counter,
                       [uint32]$failed_counter,
                       [uint32]$excluded_counter,
                       [uint32]$total_count)
{
    $time_format_string = "{0:dd\:hh\:mm\:ss\,fff}"

    $elapsed_sec = ((Get-Date) - $start_time).TotalSeconds
    $elapsed_ts =  [timespan]::fromseconds($elapsed_sec)
    $elapsed_time = ($time_format_string -f $elapsed_ts)

    if ($counter -ne 0) {
        $avg_elapsed_sec_per_file = $elapsed_sec / $counter
    } else {
        $avg_elapsed_sec_per_file = [double]::PositiveInfinity
    }
    $files_left = $total_count - $counter
    $left_sec = $files_left * $avg_elapsed_sec_per_file
    if ([double]::IsInfinity($left_sec)) {
        $estimated_time_left = $left_sec
    } else {
        $estimate_ts =  [timespan]::fromseconds($left_sec)
        $estimated_time_left = ($time_format_string -f $estimate_ts)
    }

    $status_format_string = "{0}% (done: {1}, failed: {2}, excluded: {3}, total: {4}," `
                          + " elapsed: {5}, left: {6})"
    $status = $status_format_string -f $percent_string,
                                       $done_counter,
                                       $failed_counter,
                                       $excluded_counter,
                                       $total_count,
                                       $elapsed_time,
                                       $estimated_time_left

    return $status, $files_left
}

$global:processed_count = 0
$global:percent = 0
$global:fixed_width_rounded_percent = "0.00"
$global:failed_count = 0
$global:done_count = 0
$global:excluded_count = 0

function update_global_state()
{
    $global:percent = ($global:processed_count / $global:number_of_files) * 100
    $global:fixed_width_rounded_percent = "{0:n2}" -f $global:percent
    $global:failed_count = $Error.Count
    $global:done_count = $global:processed_count - $global:failed_count
}

function is_excluded([String]$file_name,
                     [String]$exclude_pattern,
                     [String]$include_pattern)
{
    if ($exclude_pattern -ne $null -and $file_name -match $exclude_pattern) {
        $exclude = $true
    } elseif ($include_pattern -ne $null -and -not ($file_name -match $include_pattern)) {
        $exclude = $true
    } else {
        $exclude = $false
    }
    return $exclude
}

try {
    $Error.clear()
    $jobs_started = 0
    $start_time = Get-Date
    foreach ($f in $CloudFiles) {
        $excluded = is_excluded $f $ExcludePattern $IncludePattern
        if ($excluded) {
            ++$global:excluded_count
            ++$global:processed_count
        }

        update_global_state

        if ($global:done_count -lt 0) {
            # account for leftover jobs, if the script was interrupted
            $global:done_count = 0
            $Error.clear()
        }

        $status, $files_left = update_status $global:processed_count `
                                             $start_time `
                                             $global:fixed_width_rounded_percent `
                                             $global:done_count `
                                             $global:failed_count `
                                             $global:excluded_count `
                                             $number_of_files
        $file = $f.FullName

        if ($excluded) {
            $message = "Excluding $file"
        } else {
            $message = "Syncing $file..."
        }
        Write-Progress -Activity $message -Status $status -PercentComplete $global:percent

        if ($excluded -and ($files_left -ne $jobs_started)) {
            continue
        }

        $command = {
            param($Python, $OdriveScript, $file)
            & $Python $OdriveScript sync $file
        }

        if ($global:multi_threaded) {
            if (!$excluded) {
                Start-Job $command -ArgumentList $Python, $OdriveScript, $file >$null
                ++$jobs_started
            }

            if (($jobs_started -eq $JobCount) -or ($files_left -eq $jobs_started)) {
                While (Get-Job -State "Running") {
                    $match = Get-Job | Select State | Select-String "Running"
                    $running = $match.count
                    $finished_jobs = $jobs_started - $running
                    $processed_files = $global:processed_count + $finished_jobs
                    if ($processed_files -lt 0) {
                        # account for leftover jobs, if the script was interrupted
                        $processed_files = 0
                    }
                    $status, $files_left = update_status $processed_files `
                                                         $start_time `
                                                         $global:fixed_width_rounded_percent `
                                                         $global:done_count `
                                                         $global:failed_count `
                                                         $global:excluded_count `
                                                         $number_of_files
                    Write-Progress -Activity "Waiting for $running job(s) to finish..." `
                                   -Status $status `
                                   -PercentComplete $global:percent
                    Start-Sleep -Milliseconds 100
                }

                Get-Job | Receive-Job >$null 2>&1

                if ($LiveErrorReporting) {
                    $Error > $ErrorFile
                }

                Remove-Job *
                $global:processed_count += $jobs_started
                $jobs_started = 0
            }
        } else {
            & $command $Python $OdriveScript $file >$null 2>&1
            ++$global:processed_count;
        }
    }
} finally {
    try {
        update_global_state

        $status, $files_left = update_status $global:processed_count `
                                             $start_time `
                                             $global:fixed_width_rounded_percent `
                                             $global:done_count `
                                             $global:failed_count `
                                             $global:excluded_count `
                                             $number_of_files
        $sleep_time_in_sec_to_workaround_rubbish_write_progress = 1
        Start-Sleep -Seconds $sleep_time_in_sec_to_workaround_rubbish_write_progress
        Write-Progress -Activity "Done" -Status $status -PercentComplete $global:percent
        Write-Host $status

        if ($global:multi_threaded) {
            Write-Host "Removing finished jobs... " -NoNewLine
            Remove-Job * >$null 2>&1
        }
    } finally {
        if ($global:multi_threaded) {
            Write-Host "Done"
        }

        Write-Host "Capturing errors... " -NoNewLine
        Get-Job | Receive-Job >$null 2>&1
        $Error > $ErrorFile
        Write-Host "Done"

        if ($Error.Count -ne 0) {
            Write-Host "Errors were encountered (see $ErrorFile for details)"
        }
    }
}
