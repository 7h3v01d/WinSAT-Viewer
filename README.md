# WinSAT Viewer (Win32_WinSAT GUI)

A lightweight Windows 10 GUI utility that queries Win32_WinSAT via PowerShell and displays Windows Experience Index component scores in a clean desktop interface.

This tool provides:

- Base score (WinSPRLevel)
- CPU score
- Memory score
- Disk score
- Graphics score
- Direct3D score
- Assessment state
- Assessment timestamp
- Ability to trigger winsat formal
- Raw JSON output viewer
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

1. Resolves the absolute path to powershell.exe
2. Executes:
```code
Get-CimInstance -ClassName Win32_WinSAT -Namespace root\cimv2
```
3. Converts output to JSON
4. Displays structured scores in a Tkinter interface
5. Optionally runs:
```code
winsat formal
```
to refresh system assessment

PowerShell path resolution handles:

- Standard 64-bit path
- Sysnative fallback (for 32-bit Python on 64-bit Windows)
- PATH fallback

## 🚀 Running From Source
```bash
python winsat_gui.py
```

---

## 📦 Building a Portable Single-File EXE
Install PyInstaller
```bash
python -m pip install pyinstaller
```
### Build
```bash
pyinstaller --onefile --windowed --clean --name WinSATViewer winsat_gui.py
```
Output:
```code
dist/WinSATViewer.exe
```
The resulting EXE is portable and does not require Python installed.

## 🔍 Example Output

Example WinSAT result:
```code
CPUScore              : 8.9
MemoryScore           : 8.9
DiskScore             : 8.2
GraphicsScore         : 6.6
D3DScore              : 9.9
WinSPRLevel           : 6.6
WinSATAssessmentState : 1 (Valid/Completed)
```
The base score equals the lowest subscore.

## ⚠️ Running WinSAT Assessment

The “Run WinSAT Assessment” button executes:
```code
winsat formal
```
This may require Administrator privileges.

If it fails:

- Run the EXE from an elevated (Admin) terminal.

## 🛠 Troubleshooting
## PowerShell Not Found

The application resolves the absolute PowerShell path automatically.<br>
If it still fails, verify:

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

exists.

## No Win32_WinSAT Data Returned

Run manually:
```code
winsat formal
```
Then refresh.

## Antivirus Flags the EXE

Some AV engines flag PyInstaller single-file builds.

If this occurs:

- Build with --onedir instead of --onefile
- Or sign the executable

## 📁 Project Structure
```code
winsat_gui.py
README.md
```
After build:
```code
dist/
    WinSATViewer.exe
```
## 📈 Potential Future Enhancements

- Automatic bottleneck detection (lowest score highlight)
- CSV/JSON export
- System hardware snapshot panel
- Direct WMI querying via pywin32 (remove PowerShell dependency)
- Modern PySide6 UI variant
- Signed enterprise build pipeline

## 📄 License

This project is provided as-is for educational and utility purposes.
