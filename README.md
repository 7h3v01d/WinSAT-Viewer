# WinSAT Viewer

<img width="689" height="537" alt="WinSAT" src="https://github.com/user-attachments/assets/bb6e283a-a8ba-40e2-b597-1cd7052df687" /><br>

A lightweight Windows 10 GUI utility that queries `Win32_WinSAT` via PowerShell and displays Windows Experience Index component scores in a dark, animated desktop interface.

**Features:**

- Animated arc gauge for the WinSPR base score
- Animated component score bars (CPU, Memory, Disk, Graphics, D3D)
- Assessment state and timestamp display
- Trigger `winsat formal` from the UI
- Coloured output log (errors, warnings, JSON)
- Copy raw JSON to clipboard
- Single-file portable EXE support

## 📌 Purpose

Windows 10 still maintains WinSAT performance metrics internally via WMI/CIM, but Microsoft removed the graphical Windows Experience Index UI years ago.

This project restores that visibility through a modern Python GUI.

## 🖥️ System Requirements

- Windows 10
- PowerShell (default on Windows 10)
- Python 3.10+ (for source version)

No external Python dependencies required beyond the standard library.

## ⚙️ How It Works

The application:

1. Resolves the absolute path to `powershell.exe`
2. Executes:
```powershell
Get-CimInstance -ClassName Win32_WinSAT -Namespace root\cimv2
```
3. Converts output to JSON
4. Displays structured scores with animated bars and a circular base score gauge
5. Optionally runs:
```powershell
winsat formal
```
to refresh the system assessment

PowerShell path resolution handles:

- Standard 64-bit path
- Sysnative fallback (for 32-bit Python on 64-bit Windows)
- PATH fallback

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
