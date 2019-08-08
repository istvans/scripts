param(
    [Parameter(ValueFromPipeline)][System.Array] $CloudFiles,
    [String] $OdrivePath = "D:\odrive",
    [String] $OdriveScript = "$HOME\.odrive-agent\bin\odrive.py",
    [String] $Python = "python",
    [String] $ErrorFile = "errors.txt",
    [uint32] $JobCount = 4,
    [bool] $LiveErrorReporting = $False
)

if ($JobCount -eq 0) {
    throw "JobCount cannot be zero"
}

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

$n = $CloudFiles.length
Write-Host "Found $n cloud files to sync"

$i = 0
$Error.clear()
$jobs_started = 0
$status_format_string = "{0}% (running: {1}, done: {2}, failed: {3}, total: {4}, estimate: {5})"
$avg_elapsed_sec_per_file = [double]::PositiveInfinity
$start_time = Get-Date
try {
    foreach ($f in $CloudFiles) {
        $percent = ($i / $n) * 100
        $fixed_width_rounded_percent = "{0:n2}" -f $percent
        $failed_count = $Error.Count
        $done_count = $i - $failed_count
        $files_left = $n - $i
        $left_sec = $files_left * $avg_elapsed_sec_per_file
        if ([double]::IsInfinity($left_sec)) {
            $estimate = $left_sec
        } else {
            $ts =  [timespan]::fromseconds($left_sec)
            $estimate = ("{0:hh\:mm\:ss\,fff}" -f $ts)
        }
        $status = $status_format_string -f $fixed_width_rounded_percent,
                                           0,
                                           $done_count,
                                           $failed_count,
                                           $n,
                                           $estimate
        $file = $f.FullName
        Write-Progress -Activity "Syncing $file..." -Status $status -PercentComplete $percent

        $command = {
            param($Python, $OdriveScript, $file)
            & $Python $OdriveScript sync $file
        }

        Start-Job $command -ArgumentList $Python, $OdriveScript, $file >$null
        ++$jobs_started

        if ($jobs_started -eq $JobCount -or $files_left -eq 0) {
            While (Get-Job -State "Running") {
                $match = Get-Job | Select State | Select-String "Running"
                $running = $match.count
                $status = $status_format_string -f $fixed_width_rounded_percent,
                                                   $running,
                                                   $done_count,
                                                   $failed_count,
                                                   $n,
                                                   $estimate
                Write-Progress -Activity "Waiting for $jobs_started jobs to finish..." -Status $status -PercentComplete $percent
                Start-Sleep 1
            }

            Get-Job | Receive-Job >$null 2>&1

            if ($LiveErrorReporting -eq $True) {
                $Error > $ErrorFile
            }

            Remove-Job *
            $i = $i + $jobs_started
            $jobs_started = 0
        }
        $elapsed_sec = ((Get-Date) - $start_time).TotalSeconds

        if ($i -ne 0) {
            $avg_elapsed_sec_per_file = $elapsed_sec / $i
        }
    }
} finally {
    Write-Host "Interrupted. Removing finished jobs... " -NoNewLine

    Get-Job | Receive-Job >$null 2>&1
    $Error > $ErrorFile

    Remove-Job *
    Write-Host "Done"
    if ($Error.Count -ne 0) {
        Write-Host "Errors were encountered (see $ErrorFile for details)"
    }
}
