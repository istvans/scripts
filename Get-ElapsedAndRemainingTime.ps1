function Get-ElapsedAndRemainingTime {
    param(
        [Parameter(Mandatory)]
        [DateTime]$StartTime,
        [Parameter(Mandatory)]
        [uint32]$ProcessedCount,
        [Parameter(Mandatory)]
        [uint32]$TotalCount,
        [string]$TimeFormat = "{0:dd\:hh\:mm\:ss\,fff}"
    )

    $now = Get-Date
    $elapsedSeconds = ($now - $StartTime).TotalSeconds
    $elapsedTimeSpan =  [TimeSpan]::fromseconds($elapsedSeconds)
    $elapsedTime = ($TimeFormat -f $elapsedTimeSpan)

    if ($TotalCount -eq 0) {
        $averageElapsedSecondsPerThing = [double]::PositiveInfinity
    }
    else {
        $averageElapsedSecondsPerThing = $elapsedSeconds / $ProcessedCount
    }

    $thingsToProcessCount = $TotalCount - $ProcessedCount

    $remainingSeconds = $thingsToProcessCount * $averageElapsedSecondsPerThing
    if ([double]::IsInfinity($remainingSeconds)) {
        $estimatedRemainingTime = $remainingSeconds
    } else {
        $estimateRemainingTimeSpan =  [TimeSpan]::fromseconds($remainingSeconds)
        $estimatedRemainingTime = ($TimeFormat -f $estimateRemainingTimeSpan)
    }

    @{
        ElapsedTime = $elapsedTime
        RemainingTime = $estimatedRemainingTime
    }
}
