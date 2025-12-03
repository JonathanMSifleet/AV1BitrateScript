param (
    [string]$folderPath = $null
)

# Use passed folder path or current location
if (-not $folderPath) {
    $folderPath = Get-Location
}

Write-Host "Scanning folder: $folderPath`n"

# Get all files in the directory and check filenames manually
$videoFiles = Get-ChildItem -Path $folderPath | Where-Object {
    $_.PSIsContainer -eq $false -and $_.Name -match "\.(mkv|mp4)$"
}

# Check if any video files were found
if ($videoFiles.Count -eq 0) {
    Write-Host "No MKV or MP4 files found in $folderPath"
    Read-Host -Prompt "Press Enter to exit"
    return
}

$totalBitrate = 0
$fileCount = 0
$useFallback = $false

foreach ($file in $videoFiles) {
    $filePath = $file.FullName
    Write-Host "Processing: $filePath"

    # Extract the extension (last 3 characters after the last '.')
    $extension = $file.Name.Split('.')[-1]
    $ext3 = if ($extension.Length -ge 3) { $extension.Substring($extension.Length - 3) } else { $extension }
    Write-Host "Detected Extension (Last 3 chars): $ext3"

    try {
        $ffprobeJson = & ffprobe -v quiet -print_format json -select_streams v:0 `
                        -show_entries stream=duration,codec_type,stream_size "$filePath" | ConvertFrom-Json

        $videoStream = $ffprobeJson.streams | Where-Object { $_.codec_type -eq "video" }

        if ($videoStream -and $videoStream.duration -and $videoStream.stream_size) {
            $duration = [double]$videoStream.duration
            $streamSize = [double]$videoStream.stream_size
            $bitrate = ($streamSize * 8) / $duration
            $totalBitrate += $bitrate
            $fileCount++
            Write-Host "File: $($file.Name) - Video Bitrate: $([math]::Round($bitrate / 1000)) kbps"
        } else {
            throw "No valid video stream data"
        }
    } catch {
        $useFallback = $true
        $duration = & ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "$filePath"
        $fileSizeBytes = $file.Length
        if ($duration -match '^\d+(\.\d+)?$' -and $duration -gt 0) {
            $durationSec = [double]$duration
            $bitrate = ($fileSizeBytes * 8) / $durationSec
            $totalBitrate += $bitrate
            $fileCount++
            Write-Host "File: $($file.Name) - Fallback Bitrate: $([math]::Round($bitrate / 1000)) kbps"
        } else {
            Write-Host "File: $($file.Name) - Error calculating duration or file size"
        }
    }
}

if ($fileCount -gt 0) {
    $multiplier = if ($useFallback) { 0.7 } else { 0.75 }
    $averageBitrate = $totalBitrate / $fileCount
    $adjustedBitrate = [math]::Round(($averageBitrate * $multiplier) / 1000)
    Write-Host "`nAdjusted Average Bitrate: $adjustedBitrate kbps across $fileCount files"
} else {
    Write-Host "No valid video bitrate data found."
}

Read-Host -Prompt "Press Enter to exit"