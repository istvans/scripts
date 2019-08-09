param(
    [Parameter(ValueFromPipeline)][System.Array] $CloudFiles,
    [String] $OdrivePath = "D:\odrive",
    [String] $OdriveScript = "$HOME\.odrive-agent\bin\odrive.py",
    [String] $Python = "python",
    [String] $ErrorFile = "errors.txt",
    [uint32] $JobCount = 8,
    [switch] $LiveErrorReporting,
    [switch] $JustFind
)

$multi_threaded = if ($JobCount -ne 0) {$True} else {$False}
$mt_status = if ($multi_threaded) {"enabled"} else {"disabled"}
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

if ($CloudFiles.length -eq 0) {
    Write-Host "Gathering cloud files in '$OdrivePath'... Please wait... " -NoNewLine
    $CloudFiles = ls -Include *.cloud* -Path $OdrivePath -Recurse
}

$number_of_files = $CloudFiles.length
Write-Host "Found $number_of_files cloud files to sync"

if ($JustFind) {
    Write-Host "JustFind was specified, so just returning the found files"
    return $CloudFiles
}

function update_status([uint32]$processed_count,
                       [DateTime]$start_time,
                       [string]$fixed_width_rounded_percent,
                       [uint32]$done_count,
                       [uint32]$failed_count,
                       [uint32]$total_count)
{
    $status_format_string = "{0}% (done: {1}, failed: {2}, total: {3}," `
                          + " elapsed: {4}, left: {5})"
    $time_format_string = "{0:dd\:hh\:mm\:ss\,fff}"

    $elapsed_sec = ((Get-Date) - $start_time).TotalSeconds
    $elapsed_ts =  [timespan]::fromseconds($elapsed_sec)
    $elapsed_time = ($time_format_string -f $elapsed_ts)

    if ($processed_count -ne 0) {
        $avg_elapsed_sec_per_file = $elapsed_sec / $processed_count
    } else {
        $avg_elapsed_sec_per_file = [double]::PositiveInfinity
    }
    $files_left = $total_count - $processed_count
    $left_sec = $files_left * $avg_elapsed_sec_per_file
    if ([double]::IsInfinity($left_sec)) {
        $estimated_time_left = $left_sec
    } else {
        $estimate_ts =  [timespan]::fromseconds($left_sec)
        $estimated_time_left = ($time_format_string -f $estimate_ts)
    }

    $status = $status_format_string -f $fixed_width_rounded_percent,
                                       $done_count,
                                       $failed_count,
                                       $total_count,
                                       $elapsed_time,
                                       $estimated_time_left

    return $status, $files_left
}

try {
    $Error.clear()
    $processed_count = 0
    $jobs_started = 0
    $start_time = Get-Date
    foreach ($f in $CloudFiles) {
        $percent = ($processed_count / $number_of_files) * 100
        $fixed_width_rounded_percent = "{0:n2}" -f $percent
        $failed_count = $Error.Count
        $done_count = $processed_count - $failed_count
        if ($done_count -lt 0) {
            # account for leftover jobs, if the script was interrupted
            $done_count = 0
            $Error.clear()
        }

        $status, $files_left = update_status $processed_count `
                                             $start_time `
                                             $fixed_width_rounded_percent `
                                             $done_count `
                                             $failed_count `
                                             $number_of_files
        $file = $f.FullName
        Write-Progress -Activity "Syncing $file..." -Status $status -PercentComplete $percent

        $command = {
            param($Python, $OdriveScript, $file)
            & $Python $OdriveScript sync $file
        }

        if ($multi_threaded) {
            Start-Job $command -ArgumentList $Python, $OdriveScript, $file >$null
            ++$jobs_started

            if ($jobs_started -eq $JobCount -or $files_left -eq 0) {
                While (Get-Job -State "Running") {
                    $match = Get-Job | Select State | Select-String "Running"
                    $running = $match.count
                    $finished_jobs = $jobs_started - $running
                    $processed_files = $processed_count + $finished_jobs
                    if ($processed_files -lt 0) {
                        # account for leftover jobs, if the script was interrupted
                        $processed_files = 0
                    }
                    $status, $files_left = update_status $processed_files `
                                                         $start_time `
                                                         $fixed_width_rounded_percent `
                                                         $done_count `
                                                         $failed_count `
                                                         $number_of_files
                    Write-Progress -Activity "Waiting for $running job(s) to finish..." -Status $status -PercentComplete $percent
                    Start-Sleep -Milliseconds 100
                }

                Get-Job | Receive-Job >$null 2>&1

                if ($LiveErrorReporting) {
                    $Error > $ErrorFile
                }

                Remove-Job *
                $processed_count += $jobs_started
                $jobs_started = 0
            }
        } else {
            & $command $Python $OdriveScript $file >$null 2>&1
            ++$processed_count;
        }
    }
} finally {
    try {
        if ($multi_threaded) {
            Write-Host "Removing finished jobs... " -NoNewLine
            Remove-Job *
        }
    } finally {
        if ($multi_threaded) {
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
