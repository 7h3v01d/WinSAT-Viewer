# WinSAT Viewer

A lightweight Windows 10 GUI utility that queries `Win32_WinSAT` via PowerShell and displays Windows Experience Index component scores in a dark, animated desktop interface.

<img width="862" height="672" alt="WinSAT" src="https://github.com/user-attachments/assets/7037fc1a-c73f-4a19-a425-e78d966a37d2" />


### Features:

- Dark, GitHub-inspired theme with a teal / amber accent palette
- Animated circular arc gauge for the WinSPR base score
- Animated, glowing component score bars (CPU, Memory, Disk, Graphics, D3D)
- Live status pill (Ready / Working / etc.) in the header
- Assessment state and last-run timestamp display
- Auto-refreshes scores on launch
- Trigger `winsat formal` from the UI
- Colour-coded output log (success, warnings, errors, raw JSON)
- Copy raw JSON to clipboard
- Single-file portable EXE support

## 📌 Purpose

Windows 10 still maintains WinSAT performance metrics internally via WMI/CIM, but Microsoft removed the graphical Windows Experience Index UI years ago.

This project restores that visibility through a modern Python GUI.

## 🖥️ System Requirements

- Windows 10
- PowerShell (default on Windows 10)
- Python 3.10+ (for source version)

No external Python dependencies required — the UI is built entirely on the standard library (`tkinter`).

## ⚙️ How It Works

The application:

1. Resolves the absolute path to `powershell.exe`
2. Executes:
```powershell
Get-CimInstance -ClassName Win32_WinSAT -Namespace root\cimv2
```
3. Converts output to JSON
4. Displays the structured scores in the UI — an animated circular gauge for the base score, animated glowing bars for each component, and a meta panel with assessment state and last-run time
5. Automatically runs a refresh on startup, and can be re-triggered anytime via the **Refresh Scores** button
6. Optionally runs:
```powershell
winsat formal
```
to refresh the system assessment (via the **Run Assessment** button), then automatically re-queries and updates the UI

PowerShell path resolution handles:

- Standard 64-bit path
- Sysnative fallback (for 32-bit Python on 64-bit Windows)
- PATH fallback

## 🎨 Interface Overview

- **Header** — app title and a status pill showing the current state (Ready, Querying…, Running assessment…, etc.)
- **Base Score Gauge** — a circular animated arc showing the WinSPR base score (the lowest of the component scores)
- **Meta Panel** — assessment state (e.g. *Valid / Completed*) and the timestamp of the last WinSAT run
- **Component Scores** — animated bars for CPU, Memory (RAM), Disk (SSD), Graphics, and D3D Gaming, each easing into place with a soft glow
- **Action Buttons**:
  - **⟳ Refresh Scores** — re-queries `Win32_WinSAT` and updates the display
  - **▶ Run Assessment** — runs `winsat formal`, then auto-refreshes
  - **⎘ Copy JSON** — copies the raw JSON response to the clipboard
- **Output Log** — a colour-coded panel on the right showing query status, raw JSON, warnings, and errors as they happen

## 🚀 Running From Source

```bash
python WinSAT_Viewer.py
```

---

## 📦 Building a Portable Single-File EXE

Install PyInstaller:

```bash
python -m pip install pyinstaller
```

Build:

```bash
pyinstaller --onefile --windowed --clean --name WinSATViewer WinSAT_Viewer.py
```

Output:

```
dist/WinSATViewer.exe
```

The resulting EXE is portable and does not require Python installed.

## 🔍 Example Output

```
CPUScore              : 8.9
MemoryScore           : 8.9
DiskScore             : 8.2
GraphicsScore         : 6.6
D3DScore              : 9.9
WinSPRLevel           : 6.6
WinSATAssessmentState : Valid / Completed
```

The base score equals the lowest subscore.

## ⚠️ Running WinSAT Assessment

The **Run Assessment** button executes `winsat formal`.

This may require Administrator privileges. If it fails, run the EXE from an elevated (Admin) terminal.

## 🛠 Troubleshooting

**PowerShell Not Found**

The application resolves the absolute PowerShell path automatically. If it still fails, verify that this exists:

```
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
```

**No Win32_WinSAT Data Returned**

Run `winsat formal` manually from an elevated terminal, then click Refresh.

**Antivirus Flags the EXE**

Some AV engines flag PyInstaller single-file builds. If this occurs:

- Build with `--onedir` instead of `--onefile`
- Or sign the executable

## 📁 Project Structure

```
WinSAT_Viewer.py
README.md
```

After build:

```
dist/
    WinSATViewer.exe
```

## 📈 Potential Future Enhancements

- Automatic bottleneck detection (lowest score highlight)
- CSV / JSON export
- System hardware snapshot panel
- Direct WMI querying via `pywin32` (remove PowerShell dependency)
- Score history tracking with trend graphs
- Signed enterprise build pipeline

## 📄 License

Copyright 2024 Leon Priest  
GitHub: [7h3v01d](https://github.com/7h3v01d)

Licensed under the [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0).
