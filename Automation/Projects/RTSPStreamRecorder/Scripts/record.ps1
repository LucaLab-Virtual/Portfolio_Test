$ffmpeg = "C:\ffmpeg-8.1.1-full_build\ffmpeg-8.1.1-full_build\bin\ffmpeg.exe"
$rtsp   = "rtsp://User:Passwordg@ipaddress/stream1"
$outDir = "D:\OutdoorCamera\Recordings"

New-Item -ItemType Directory -Path $outDir -Force | Out-Null

while ($true) {
    try {
        & $ffmpeg `
        -rtsp_transport tcp `
        -use_wallclock_as_timestamps 1 `
        -fflags +genpts `
        -i $rtsp `
        -map 0 `
        -c:v copy `
        -c:a aac `
        -b:a 128k `
        -f segment `
        -segment_time 900 `
        -reset_timestamps 1 `
        -strftime 1 `
        "$outDir\camera_%Y-%m-%d_%H-%M-%S.ts"

        Start-Sleep -Seconds 5
    }
    catch {
        Start-Sleep -Seconds 5
    }
}
