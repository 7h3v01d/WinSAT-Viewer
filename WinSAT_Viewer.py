import json
import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox


APP_TITLE = "WinSAT (Win32_WinSAT) Viewer — Windows 10"


def resolve_powershell_path() -> str:
    """
    Resolve a reliable path to Windows PowerShell on Windows 10.
    Handles PATH issues and 32-bit Python on 64-bit Windows (Sysnative).
    """
    system_root = os.environ.get("SystemRoot", r"C:\Windows")

    # Standard location (may be redirected under 32-bit Python on 64-bit Windows)
    ps_system32 = os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")

    # Sysnative bypasses file system redirection for 32-bit processes on 64-bit Windows
    ps_sysnative = os.path.join(system_root, "Sysnative", "WindowsPowerShell", "v1.0", "powershell.exe")

    if os.path.exists(ps_system32):
        return ps_system32
    if os.path.exists(ps_sysnative):
        return ps_sysnative

    # Fallback to PATH resolution
    return "powershell"


def run_powershell(ps_script: str, timeout_sec: int = 30) -> tuple[int, str, str]:
    """
    Run PowerShell with a script and return (returncode, stdout, stderr).
    Uses -NoProfile and Bypass policy for reliability.

    NOTE: Uses resolved absolute path to avoid PATH-related failures.
    """
    ps_exe = resolve_powershell_path()
    cmd = [
        ps_exe,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        ps_script,
    ]
    try:
        p = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout_sec,
            encoding="utf-8",
            errors="replace",
        )
        return p.returncode, p.stdout.strip(), p.stderr.strip()
    except FileNotFoundError:
        return 127, "", f"PowerShell executable not found at: {ps_exe}"
    except subprocess.TimeoutExpired:
        return 124, "", f"PowerShell timed out after {timeout_sec} seconds."


def ps_query_winsat_json() -> str:
    # Note: Win10 commonly exposes WinSATAssessmentState (not AssessmentState).
    return r"""
$ErrorActionPreference = "Stop"
function Get-WinsatObj {
  try { return Get-CimInstance -ClassName Win32_WinSAT -Namespace root\cimv2 }
  catch { return Get-WmiObject -Class Win32_WinSAT -Namespace root\cimv2 }
}
$w = Get-WinsatObj
if (-not $w) { throw "No Win32_WinSAT instance returned. WinSAT may not be available on this system." }

$wPicked = $w | Sort-Object -Property TimeTaken -Descending -ErrorAction SilentlyContinue | Select-Object -First 1

$wPicked |
  Select-Object CPUScore, D3DScore, DiskScore, MemoryScore, GraphicsScore, WinSPRLevel, TimeTaken, WinSATAssessmentState |
  ConvertTo-Json -Depth 3
"""


def ps_run_winsat_assessment() -> str:
    return r"""
$ErrorActionPreference = "Stop"
winsat formal | Out-String
"""


def decode_assessment_state(v) -> str:
    """
    WinSATAssessmentState is commonly:
      1 = Valid/Completed
      0 = Not run/Unknown
    Meanings can vary; keep conservative labeling.
    """
    try:
        iv = int(v)
    except Exception:
        return str(v)

    if iv == 1:
        return "1 (Valid/Completed)"
    if iv == 0:
        return "0 (Not run/Unknown)"
    return f"{iv} (Unknown meaning)"


class WinSatGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x560")
        self.minsize(720, 480)

        self.raw_json: str | None = None

        self._build_ui()
        self._set_busy(False)

        self.after(200, self.refresh_scores)

    def _build_ui(self):
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x")

        ttk.Label(header, text="Win32_WinSAT Scores", font=("Segoe UI", 16, "bold")).pack(side="left")

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(header, textvariable=self.status_var).pack(side="right")

        btn_row = ttk.Frame(outer)
        btn_row.pack(fill="x", pady=(10, 8))

        self.refresh_btn = ttk.Button(btn_row, text="Refresh (Query Win32_WinSAT)", command=self.refresh_scores)
        self.refresh_btn.pack(side="left")

        self.run_btn = ttk.Button(btn_row, text="Run WinSAT Assessment (winsat formal)", command=self.run_assessment)
        self.run_btn.pack(side="left", padx=(8, 0))

        self.copy_btn = ttk.Button(btn_row, text="Copy JSON", command=self.copy_json)
        self.copy_btn.pack(side="left", padx=(8, 0))

        content = ttk.PanedWindow(outer, orient="vertical")
        content.pack(fill="both", expand=True)

        score_frame = ttk.Labelframe(content, text="Scores", padding=10)
        content.add(score_frame, weight=2)

        grid = ttk.Frame(score_frame)
        grid.pack(fill="x", expand=False)

        self.fields = {
            "WinSPRLevel (Base Score)": tk.StringVar(value="-"),
            "CPUScore": tk.StringVar(value="-"),
            "MemoryScore": tk.StringVar(value="-"),
            "DiskScore": tk.StringVar(value="-"),
            "GraphicsScore": tk.StringVar(value="-"),
            "D3DScore": tk.StringVar(value="-"),
            "WinSATAssessmentState": tk.StringVar(value="-"),
            "TimeTaken": tk.StringVar(value="-"),
        }

        for r, (label, var) in enumerate(self.fields.items()):
            ttk.Label(grid, text=f"{label}:", width=30).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Label(grid, textvariable=var, font=("Consolas", 11)).grid(row=r, column=1, sticky="w", pady=3)

        lower = ttk.Labelframe(content, text="Raw JSON / Log", padding=8)
        content.add(lower, weight=3)

        self.text = tk.Text(lower, height=14, wrap="word")
        self.text.pack(fill="both", expand=True)
        self.text.configure(state="disabled")

        # Diagnostics footer (shows the resolved PowerShell path)
        diag = ttk.Frame(outer)
        diag.pack(fill="x", pady=(8, 0))
        self.ps_path_var = tk.StringVar(value=f"PowerShell: {resolve_powershell_path()}")
        ttk.Label(diag, textvariable=self.ps_path_var).pack(side="left")

    def _set_busy(self, busy: bool, msg: str = ""):
        if busy:
            self.status_var.set(msg or "Working…")
            self.refresh_btn.configure(state="disabled")
            self.run_btn.configure(state="disabled")
            self.copy_btn.configure(state="disabled")
            self.configure(cursor="watch")
        else:
            self.status_var.set(msg or "Ready.")
            self.refresh_btn.configure(state="normal")
            self.run_btn.configure(state="normal")
            self.copy_btn.configure(state="normal")
            self.configure(cursor="")

    def _log(self, s: str):
        self.text.configure(state="normal")
        self.text.insert("end", s + "\n")
        self.text.see("end")
        self.text.configure(state="disabled")

    def _clear_log(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

    def _apply_scores(self, obj: dict):
        mapping = {
            "WinSPRLevel (Base Score)": "WinSPRLevel",
            "CPUScore": "CPUScore",
            "MemoryScore": "MemoryScore",
            "DiskScore": "DiskScore",
            "GraphicsScore": "GraphicsScore",
            "D3DScore": "D3DScore",
            "WinSATAssessmentState": "WinSATAssessmentState",
            "TimeTaken": "TimeTaken",
        }
        for ui_label, key in mapping.items():
            val = obj.get(key, None)
            if val is None:
                self.fields[ui_label].set("-")
            else:
                if key == "WinSATAssessmentState":
                    self.fields[ui_label].set(decode_assessment_state(val))
                else:
                    self.fields[ui_label].set(str(val))

    def copy_json(self):
        if not self.raw_json:
            messagebox.showinfo(APP_TITLE, "No JSON available yet. Click Refresh first.")
            return
        self.clipboard_clear()
        self.clipboard_append(self.raw_json)
        self.status_var.set("JSON copied to clipboard.")

    def refresh_scores(self):
        def worker():
            self._clear_log()
            self._log("Querying Win32_WinSAT via PowerShell…")
            rc, out, err = run_powershell(ps_query_winsat_json(), timeout_sec=30)
            self.after(0, lambda: self._handle_query_result(rc, out, err))

        self._set_busy(True, "Querying Win32_WinSAT…")
        threading.Thread(target=worker, daemon=True).start()

    def _handle_query_result(self, rc: int, out: str, err: str):
        try:
            if rc != 0:
                self._log(f"[ERROR] PowerShell exit code: {rc}")
                if err:
                    self._log(err)
                if out:
                    self._log(out)
                messagebox.showerror(APP_TITLE, "Failed to query Win32_WinSAT.\nSee log for details.")
                return

            if not out:
                self._log("[ERROR] Empty output.")
                messagebox.showerror(APP_TITLE, "PowerShell returned no output.")
                return

            self.raw_json = out
            self._log("Raw JSON:")
            self._log(out)

            obj = json.loads(out)
            if isinstance(obj, list) and obj:
                obj = obj[0]
            if not isinstance(obj, dict):
                raise ValueError("Unexpected JSON structure (not an object).")

            self._apply_scores(obj)
            self.status_var.set("WinSAT scores updated.")
        except Exception as e:
            self._log(f"[ERROR] {e}")
            messagebox.showerror(APP_TITLE, f"Could not parse/display results:\n{e}")
        finally:
            self._set_busy(False)

    def run_assessment(self):
        def worker():
            self._clear_log()
            self._log("Running: winsat formal …")
            self._log("Note: this can take a while and may require Administrator privileges.\n")
            rc, out, err = run_powershell(ps_run_winsat_assessment(), timeout_sec=600)

            def done():
                if rc != 0:
                    self._log(f"[ERROR] WinSAT exit code: {rc}")
                    if err:
                        self._log(err)
                    if out:
                        self._log(out)
                    messagebox.showerror(
                        APP_TITLE,
                        "WinSAT assessment failed.\nTry running this program from an elevated (Admin) terminal.\nSee log for details.",
                    )
                    self._set_busy(False, "Ready.")
                    return

                if out:
                    self._log("WinSAT output:")
                    self._log(out)
                if err:
                    self._log("WinSAT stderr:")
                    self._log(err)

                self._log("\nRe-querying Win32_WinSAT…\n")
                self._set_busy(True, "Re-querying Win32_WinSAT…")
                threading.Thread(target=self._refresh_after_assessment, daemon=True).start()

            self.after(0, done)

        self._set_busy(True, "Running WinSAT formal…")
        threading.Thread(target=worker, daemon=True).start()

    def _refresh_after_assessment(self):
        rc, out, err = run_powershell(ps_query_winsat_json(), timeout_sec=30)
        self.after(0, lambda: self._handle_query_result(rc, out, err))


def main():
    app = WinSatGUI()
    app.mainloop()


if __name__ == "__main__":
    main()