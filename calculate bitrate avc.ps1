param (
    [string]$folderPath = $null
)

if (-not $folderPath) {
    $folderPath = Get-Location
}

Write-Host "Scanning folder: $folderPath`n"

$videoFiles = Get-ChildItem -Path $folderPath | Where-Object { $_.PSIsContainer -eq $false -and $_.Name -match "\.(mkv|mp4)$" }

if ($videoFiles.Count -eq 0) {
    Write-Host "No MKV or MP4 files found in $folderPath"
    Read-Host -Prompt "Press Enter to exit"
    return
}

$totalBitrate = 0
$fileCount = 0

foreach ($file in $videoFiles) {
    Write-Host "---------------------------------------------"
    $filePath = $file.FullName
    Write-Host "Processing: $filePath"

    $extension = $file.Name.Split('.')[-1]
    $ext3 = if ($extension.Length -ge 3) { $extension.Substring($extension.Length - 3) } else { $extension }
    Write-Host "Detected Extension (Last 3 chars): $ext3"

    $audioBitrateTotal = 0

    # Get MediaInfo data
    $mediaInfoJson = & MediaInfo --Output=JSON "$filePath" | ConvertFrom-Json
    $tracks = $mediaInfoJson.media.track

    # Overall bitrate from General track
    $overallBitRate = 0
    $generalTrack = $tracks | Where-Object { $_.'@type' -eq 'General' } | Select-Object -First 1
    if ($generalTrack -and $generalTrack.OverallBitRate -and $generalTrack.OverallBitRate -match '^\d+$') {
        $overallBitRate = [double]$generalTrack.OverallBitRate
    }

    # Sum audio bitrates
    $audioTracks = $tracks | Where-Object { $_.'@type' -eq 'Audio' }
    foreach ($audio in $audioTracks) {
        if ($audio.BitRate -and $audio.BitRate -match '^\d+$') {
            $audioBitrateTotal += [double]$audio.BitRate
        }
    }

    # Video bitrate (accurate direct value)
    $videoBitrate = 0
    $videoTrack = $tracks | Where-Object { $_.'@type' -eq 'Video' } | Select-Object -First 1
    if ($videoTrack -and $videoTrack.BitRate -and $videoTrack.BitRate -match '^\d+$') {
        $videoBitrate = [double]$videoTrack.BitRate
    }

    if ($videoBitrate -gt 0) {
        $totalBitrate += $videoBitrate
        $fileCount++

        Write-Host "File: $($file.Name) - Fallback Total Bitrate: $([math]::Round($overallBitRate / 1000)) kbps"
        Write-Host "Deducted audio bitrate: $([math]::Round($audioBitrateTotal / 1000)) kbps"
        Write-Host "Video-only estimated bitrate: $([math]::Round($videoBitrate / 1000)) kbps"
    } else {
        Write-Host "File: $($file.Name) - Could not determine video bitrate"
    }
}

Write-Host "---------------------------------------------"

if ($fileCount -gt 0) {
    $multiplier = 0.375
    $averageBitrate = $totalBitrate / $fileCount
    $averageBitrateKbps = [math]::Round($averageBitrate / 1000)
    Write-Host "Average Bitrate BEFORE multiplier: $averageBitrateKbps kbps"
    $adjustedBitrate = [math]::Round(($averageBitrate * $multiplier) / 1000)
    Write-Host "Adjusted Average Bitrate: $adjustedBitrate kbps across $fileCount files"
} else {
    Write-Host "No valid video bitrate data found."
}

Read-Host -Prompt "Press Enter to exit"