# 📦 FFmpeg RTSP Stream Recorder (Lightweight NVR)

![PowerShell](https://img.shields.io/badge/PowerShell-Automation-blue)
![FFmpeg](https://img.shields.io/badge/FFmpeg-Streaming-green)
![Windows](https://img.shields.io/badge/Platform-Windows-lightgrey)
![Status](https://img.shields.io/badge/Status-Production%20Ready-success)

---

## 🎯 Overview

This project implements a **lightweight Network Video Recorder (NVR)** using FFmpeg and PowerShell.

It provides:

* 🎥 Continuous RTSP recording
* 🔁 Automatic restart on failure
* ⏱ Time-based segmentation
* 🧹 Automatic cleanup (retention policy)
* ⚙️ Background execution at system startup

---

## 🧠 Architecture

```id="zdsq2k"
        +----------------------+
        |   IP Camera (RTSP)   |
        | 192.168.18.138       |
        +----------+-----------+
                   |
                   | RTSP Stream
                   v
        +----------------------+
        |      FFmpeg          |
        |  (Processing Engine) |
        +----------+-----------+
                   |
                   | .ts Segments
                   v
        +----------------------+
        |   Local Storage      |
        | D:\OutdoorCamera     |
        +----------+-----------+
                   |
        +----------+-----------+
        | PowerShell Scripts   |
        | - record.ps1         |
        | - cleanup.ps1        |
        +----------+-----------+
                   |
        +----------+-----------+
        | Task Scheduler       |
        | (Auto Start / Clean) |
        +----------------------+
```

---

## 📂 Project Structure

```id="d0nxfu"
rtsp-stream-recorder-ffmpeg/
│
├── README.md
├── scripts/
│   ├── record.ps1
│   ├── cleanup.ps1
│   ├── CreateScheduledTaskRECORDER.ps1
│   └── CleanUpTaskDIALY.ps1
│
└── docs/
    └── guide.pdf (optional)
```

---

## 🧠 What is FFmpeg?

FFmpeg is a powerful multimedia framework used for:

* Recording RTSP streams
* Converting video formats
* Streaming media
* Processing audio/video

---

## ⬇️ Installation

### 1. Download FFmpeg

https://ffmpeg.org/download.html

Recommended:

* `ffmpeg-8.1.1-full_build`

---

### 2. Extract

```id="0r0k4j"
C:\ffmpeg\
```

---

## 📡 RTSP URL Structure

```id="ph0wfu"
rtsp://username:password@ip_address/stream
```

Example:

```id="3s5b7c"
rtsp://Lucarez1:Lucarezking@192.168.18.138/stream1
```

---

## 🎥 Recording Script

Full script available here:  
👉 [record.ps1](Scripts/record.ps1)

### What it does:
- Connects to RTSP stream
- Records continuously
- Splits video into X-minute segments
- Auto-restarts if camera disconnects

---

## 🧹 Cleanup Script

Full script:  
👉 [cleanup.ps1](Scripts/cleanup.ps1)

Deletes recordings older than X days.

---

## ⚙️ Scheduled Task (Recorder)

👉 [CreateScheduledTaskRECORDER.ps1](Scripts/CreateScheduledTaskRECORDER.ps1)

Runs the recorder automatically at system startup.

---

## ⚙️ Scheduled Task (Cleanup)

👉 [CleanUpTaskDIALY.ps1](Scripts/CleanUpTaskDIALY.ps1)

Runs daily cleanup at scheduled time.

---

## ⚙️ How It Works

1. Windows boots
2. Task Scheduler starts `record.ps1`
3. FFmpeg connects to RTSP stream
4. Video is recorded into `.ts` segments
5. If connection fails → script restarts
6. Cleanup task deletes old files daily

---

## 📁 Output

```id="r4n4n1"
D:\OutdoorCamera\
camera_2026-06-17_10-00-00.ts
camera_2026-06-17_10-10-00.ts
```

---

## ⚠️ Notes

* `.ts` format ensures stability
* No re-encoding → low CPU usage
* Works even if interrupted

---

## 🚀 Future Improvements

* Multi-camera support
* Web dashboard (live view)
* Cloud backup (S3 / Google Drive)
* Motion detection

---

## 👨‍💻 Author

Luis Carlos Ramirez
IT Support Specialist / Cloud & Systems Administrator

---
