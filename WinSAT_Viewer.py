import json
import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont


APP_TITLE = "WinSAT Viewer"
APP_SUBTITLE = "Win32_WinSAT System Assessment"

# ── Palette ──────────────────────────────────────────────────────────────────
C = {
    "bg":         "#0d1117",   # near-black base
    "surface":    "#161b22",   # card / panel surface
    "border":     "#21262d",   # subtle borders
    "border_hi":  "#30363d",   # highlighted borders
    "teal":       "#00d4aa",   # primary accent (teal)
    "teal_dim":   "#00856a",   # dimmed teal for tracks
    "amber":      "#f0a500",   # score value accent
    "amber_dim":  "#7a5400",   # amber track
    "red":        "#f85149",   # error states
    "green":      "#3fb950",   # success / active
    "text_hi":    "#e6edf3",   # high-emphasis text
    "text_med":   "#8b949e",   # medium-emphasis text
    "text_lo":    "#484f58",   # low-emphasis / disabled
    "overlay":    "#1f2937",   # tooltips / overlays
}

SCORE_BARS = {
    "CPUScore":      ("CPU",          C["teal"]),
    "MemoryScore":   ("Memory (RAM)", C["teal"]),
    "DiskScore":     ("Disk (SSD)",   C["teal"]),
    "GraphicsScore": ("Graphics",     C["amber"]),
    "D3DScore":      ("D3D Gaming",   C["amber"]),
}

MAX_SCORE = 9.9


def _blend(hex_a: str, hex_b: str, t: float) -> str:
    """Linear interpolate between two #rrggbb colours. t=0 → a, t=1 → b."""
    a = int(hex_a[1:3], 16), int(hex_a[3:5], 16), int(hex_a[5:7], 16)
    b = int(hex_b[1:3], 16), int(hex_b[3:5], 16), int(hex_b[5:7], 16)
    r = tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return "#{:02x}{:02x}{:02x}".format(*r)


# Pre-compute glow shades (blended toward bg) so Canvas never sees alpha hex
_BG = C["bg"]
C["teal_glow1"] = _blend(C["teal"],  _BG, 0.75)   # very faint (~25% teal)
C["teal_glow2"] = _blend(C["teal"],  _BG, 0.45)   # medium   (~55% teal)
C["amber_glow1"] = _blend(C["amber"], _BG, 0.75)
C["amber_glow2"] = _blend(C["amber"], _BG, 0.45)


def _glow_colors(color: str):
    """Return (faint, mid) glow shades for a given accent colour."""
    if color == C["teal"]:
        return C["teal_glow1"], C["teal_glow2"]
    if color == C["amber"]:
        return C["amber_glow1"], C["amber_glow2"]
    # fallback: compute on the fly
    return _blend(color, _BG, 0.75), _blend(color, _BG, 0.45)


# ── PowerShell helpers (unchanged logic) ─────────────────────────────────────

def resolve_powershell_path() -> str:
    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    ps_system32 = os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
    ps_sysnative = os.path.join(system_root, "Sysnative", "WindowsPowerShell", "v1.0", "powershell.exe")
    if os.path.exists(ps_system32):
        return ps_system32
    if os.path.exists(ps_sysnative):
        return ps_sysnative
    return "powershell"


def run_powershell(ps_script: str, timeout_sec: int = 30) -> tuple[int, str, str]:
    ps_exe = resolve_powershell_path()
    cmd = [ps_exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script]
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout_sec,
                           encoding="utf-8", errors="replace")
        return p.returncode, p.stdout.strip(), p.stderr.strip()
    except FileNotFoundError:
        return 127, "", f"PowerShell not found: {ps_exe}"
    except subprocess.TimeoutExpired:
        return 124, "", f"Timed out after {timeout_sec}s."


def ps_query_winsat_json() -> str:
    return r"""
$ErrorActionPreference = "Stop"
function Get-WinsatObj {
  try { return Get-CimInstance -ClassName Win32_WinSAT -Namespace root\cimv2 }
  catch { return Get-WmiObject -Class Win32_WinSAT -Namespace root\cimv2 }
}
$w = Get-WinsatObj
if (-not $w) { throw "No Win32_WinSAT instance returned." }
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
    try:
        iv = int(v)
    except Exception:
        return str(v)
    if iv == 1:
        return "Valid / Completed"
    if iv == 0:
        return "Not run / Unknown"
    return f"State {iv}"


# ── Animated score bar widget ─────────────────────────────────────────────────

class ScoreBar(tk.Frame):
    """A custom animated score bar: label | track[========] | value"""

    ANIM_STEPS = 30
    ANIM_MS    = 16   # ~60 fps

    def __init__(self, parent, label: str, color: str, **kw):
        super().__init__(parent, bg=C["surface"], **kw)
        self._color      = color
        self._target     = 0.0
        self._current    = 0.0
        self._anim_id    = None

        # label
        lbl = tk.Label(self, text=label, bg=C["surface"], fg=C["text_med"],
                       font=("Segoe UI", 9), width=14, anchor="w")
        lbl.pack(side="left", padx=(0, 10))

        # canvas track
        self._canvas = tk.Canvas(self, bg=C["bg"], height=18,
                                 highlightthickness=1, highlightbackground=C["border"])
        self._canvas.pack(side="left", fill="x", expand=True)
        self._track_id = None
        self._fill_id  = None
        self._canvas.bind("<Configure>", self._on_resize)

        # value label
        self._val_var = tk.StringVar(value="—")
        val_lbl = tk.Label(self, textvariable=self._val_var, bg=C["surface"],
                           fg=color, font=("Consolas", 11, "bold"), width=6, anchor="e")
        val_lbl.pack(side="left", padx=(10, 0))

        self._draw(0.0)

    def _on_resize(self, event):
        self._draw(self._current)

    def _draw(self, fraction: float):
        c = self._canvas
        w = c.winfo_width() or 200
        h = c.winfo_height() or 18
        c.delete("all")
        # track
        c.create_rectangle(0, 0, w, h, fill=C["bg"], outline="")
        # fill
        fill_w = max(0, int(w * min(fraction, 1.0)))
        if fill_w > 0:
            glow1, glow2 = _glow_colors(self._color)
            # glow segments: faint outer → mid → full bright core
            c.create_rectangle(0, 2, fill_w, h - 2, fill=glow1,        outline="")
            c.create_rectangle(0, 4, fill_w, h - 4, fill=glow2,        outline="")
            c.create_rectangle(0, 6, fill_w, h - 6, fill=self._color,  outline="")
            # bright leading edge
            if fill_w > 2:
                c.create_rectangle(fill_w - 2, 2, fill_w, h - 2,
                                   fill=self._color, outline="")

    def set_score(self, score: float | None):
        if score is None:
            self._val_var.set("—")
            self._animate_to(0.0)
            return
        self._val_var.set(f"{score:.1f}")
        self._animate_to(score / MAX_SCORE)

    def _animate_to(self, target: float):
        if self._anim_id:
            self.after_cancel(self._anim_id)
        self._target  = target
        self._step    = 0
        self._start   = self._current
        self._tick()

    def _tick(self):
        self._step += 1
        t = self._step / self.ANIM_STEPS
        t = t * t * (3 - 2 * t)   # smoothstep easing
        self._current = self._start + (self._target - self._start) * t
        self._draw(self._current)
        if self._step < self.ANIM_STEPS:
            self._anim_id = self.after(self.ANIM_MS, self._tick)
        else:
            self._current = self._target
            self._draw(self._current)


# ── Base score ring ───────────────────────────────────────────────────────────

class BaseScoreRing(tk.Canvas):
    """Circular gauge for the WinSPR base score."""

    def __init__(self, parent, size=120, **kw):
        super().__init__(parent, width=size, height=size,
                         bg=C["surface"], highlightthickness=0, **kw)
        self._size   = size
        self._score  = None
        self._anim_v = 0.0
        self._target = 0.0
        self._anim_id = None
        self._draw(0.0)
        self.bind("<Configure>", lambda e: self._draw(self._anim_v))

    def _draw(self, fraction: float):
        s = self._size
        self.delete("all")
        pad = 12
        x0, y0, x1, y1 = pad, pad, s - pad, s - pad

        # track arc
        self.create_arc(x0, y0, x1, y1, start=220, extent=-260,
                        style="arc", outline=C["border_hi"], width=8)

        # filled arc
        if fraction > 0:
            self.create_arc(x0, y0, x1, y1, start=220, extent=int(-260 * fraction),
                            style="arc", outline=C["teal"], width=8)

        # centre text
        cx, cy = s // 2, s // 2
        if self._score is not None:
            self.create_text(cx, cy - 8, text=f"{self._score:.1f}",
                             fill=C["teal"], font=("Consolas", 22, "bold"))
            self.create_text(cx, cy + 14, text="Base Score",
                             fill=C["text_med"], font=("Segoe UI", 8))
        else:
            self.create_text(cx, cy, text="—", fill=C["text_lo"],
                             font=("Consolas", 22, "bold"))

    def set_score(self, score: float | None):
        self._score = score
        target = (score / MAX_SCORE) if score is not None else 0.0
        self._animate_to(target)

    def _animate_to(self, target: float):
        if self._anim_id:
            self.after_cancel(self._anim_id)
        self._target = target
        self._start  = self._anim_v
        self._step   = 0
        self._tick()

    def _tick(self):
        self._step += 1
        t = self._step / 40
        t = t * t * (3 - 2 * t)
        self._anim_v = self._start + (self._target - self._start) * t
        self._draw(self._anim_v)
        if self._step < 40:
            self._anim_id = self.after(16, self._tick)
        else:
            self._anim_v = self._target
            self._draw(self._anim_v)


# ── Main application ──────────────────────────────────────────────────────────

class WinSatGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("860x640")
        self.minsize(760, 540)
        self.configure(bg=C["bg"])

        self.raw_json: str | None = None
        self._score_bars: dict[str, ScoreBar] = {}
        self._meta_vars: dict[str, tk.StringVar] = {}

        self._style_ttk()
        self._build_ui()
        self._set_busy(False)

        self.after(200, self.refresh_scores)

    # ── ttk theme override ────────────────────────────────────────────────────

    def _style_ttk(self):
        s = ttk.Style(self)
        s.theme_use("clam")

        s.configure(".",
                    background=C["bg"],
                    foreground=C["text_hi"],
                    bordercolor=C["border"],
                    troughcolor=C["surface"],
                    focuscolor=C["teal"])

        s.configure("TFrame",  background=C["bg"])
        s.configure("Surface.TFrame", background=C["surface"])

        s.configure("TLabel",
                    background=C["bg"],
                    foreground=C["text_hi"],
                    font=("Segoe UI", 9))

        s.configure("Surface.TLabel",
                    background=C["surface"],
                    foreground=C["text_hi"])

        s.configure("Dim.TLabel",
                    background=C["bg"],
                    foreground=C["text_med"],
                    font=("Segoe UI", 8))

        s.configure("Meta.TLabel",
                    background=C["surface"],
                    foreground=C["text_med"],
                    font=("Segoe UI", 8))

        s.configure("MetaVal.TLabel",
                    background=C["surface"],
                    foreground=C["text_hi"],
                    font=("Consolas", 9))

        s.configure("Teal.TButton",
                    background=C["surface"],
                    foreground=C["teal"],
                    bordercolor=C["teal_dim"],
                    focuscolor=C["teal"],
                    font=("Segoe UI", 9),
                    padding=(10, 5))
        s.map("Teal.TButton",
              background=[("active", C["overlay"]), ("disabled", C["surface"])],
              foreground=[("disabled", C["text_lo"])])

        s.configure("Amber.TButton",
                    background=C["surface"],
                    foreground=C["amber"],
                    bordercolor=C["amber_dim"],
                    focuscolor=C["amber"],
                    font=("Segoe UI", 9),
                    padding=(10, 5))
        s.map("Amber.TButton",
              background=[("active", C["overlay"]), ("disabled", C["surface"])],
              foreground=[("disabled", C["text_lo"])])

        s.configure("TPanedwindow", background=C["bg"])

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        root_pad = ttk.Frame(self, padding=16)
        root_pad.pack(fill="both", expand=True)

        # ── Header ─────────────────────────────────────────────────────────
        hdr = ttk.Frame(root_pad)
        hdr.pack(fill="x", pady=(0, 14))

        # left: title block
        title_blk = ttk.Frame(hdr)
        title_blk.pack(side="left")

        tk.Label(title_blk, text="WIN", bg=C["bg"], fg=C["teal"],
                 font=("Consolas", 22, "bold")).pack(side="left")
        tk.Label(title_blk, text="SAT", bg=C["bg"], fg=C["text_hi"],
                 font=("Consolas", 22, "bold")).pack(side="left")
        tk.Label(title_blk, text="  VIEWER", bg=C["bg"], fg=C["text_lo"],
                 font=("Consolas", 14)).pack(side="left", padx=(0, 16))
        tk.Label(title_blk, text="Win32_WinSAT System Assessment", bg=C["bg"],
                 fg=C["text_med"], font=("Segoe UI", 9)).pack(side="left")

        # right: status pill
        self._status_frame = tk.Frame(hdr, bg=C["surface"],
                                      highlightthickness=1,
                                      highlightbackground=C["border"])
        self._status_frame.pack(side="right")
        self._status_dot = tk.Label(self._status_frame, text="●", bg=C["surface"],
                                    fg=C["green"], font=("Segoe UI", 8))
        self._status_dot.pack(side="left", padx=(8, 4), pady=4)
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self._status_frame, textvariable=self.status_var, bg=C["surface"],
                 fg=C["text_med"], font=("Segoe UI", 8)).pack(side="left", padx=(0, 10), pady=4)

        # thin separator line
        sep = tk.Frame(root_pad, bg=C["border"], height=1)
        sep.pack(fill="x", pady=(0, 14))

        # ── Main body: left panel + right log ──────────────────────────────
        body = ttk.Frame(root_pad)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=C["bg"])
        left.pack(side="left", fill="both", expand=False, padx=(0, 12))
        left.configure(width=480)
        left.pack_propagate(False)

        right = tk.Frame(body, bg=C["surface"],
                         highlightthickness=1,
                         highlightbackground=C["border"])
        right.pack(side="left", fill="both", expand=True)

        # ── Left: base score ring + meta row ───────────────────────────────
        top_row = tk.Frame(left, bg=C["bg"])
        top_row.pack(fill="x", pady=(0, 12))

        self._ring = BaseScoreRing(top_row, size=130)
        self._ring.configure(bg=C["bg"])
        self._ring.pack(side="left", padx=(0, 16))

        meta_card = tk.Frame(top_row, bg=C["surface"],
                             highlightthickness=1,
                             highlightbackground=C["border"])
        meta_card.pack(side="left", fill="both", expand=True, ipady=8)

        meta_inner = tk.Frame(meta_card, bg=C["surface"])
        meta_inner.pack(fill="both", expand=True, padx=12, pady=8)

        for key, label in [("WinSATAssessmentState", "Assessment State"),
                            ("TimeTaken",             "Last Run")]:
            row = tk.Frame(meta_inner, bg=C["surface"])
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label.upper(), bg=C["surface"], fg=C["text_lo"],
                     font=("Segoe UI", 7, "bold")).pack(anchor="w")
            var = tk.StringVar(value="—")
            self._meta_vars[key] = var
            tk.Label(row, textvariable=var, bg=C["surface"], fg=C["text_hi"],
                     font=("Consolas", 10)).pack(anchor="w")

        # ── Left: score bars ───────────────────────────────────────────────
        bars_card = tk.Frame(left, bg=C["surface"],
                             highlightthickness=1,
                             highlightbackground=C["border"])
        bars_card.pack(fill="x", pady=(0, 12))

        bars_inner = tk.Frame(bars_card, bg=C["surface"])
        bars_inner.pack(fill="x", padx=12, pady=10)

        tk.Label(bars_inner, text="COMPONENT SCORES", bg=C["surface"],
                 fg=C["text_lo"], font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(0, 8))

        for key, (label, color) in SCORE_BARS.items():
            bar = ScoreBar(bars_inner, label=label, color=color)
            bar.pack(fill="x", pady=3)
            self._score_bars[key] = bar

        # ── Left: action buttons ───────────────────────────────────────────
        btn_row = tk.Frame(left, bg=C["bg"])
        btn_row.pack(fill="x")

        self.refresh_btn = ttk.Button(btn_row, text="⟳  Refresh Scores",
                                      style="Teal.TButton", command=self.refresh_scores)
        self.refresh_btn.pack(side="left")

        self.run_btn = ttk.Button(btn_row, text="▶  Run Assessment",
                                  style="Amber.TButton", command=self.run_assessment)
        self.run_btn.pack(side="left", padx=(8, 0))

        self.copy_btn = ttk.Button(btn_row, text="⎘  Copy JSON",
                                   style="Teal.TButton", command=self.copy_json)
        self.copy_btn.pack(side="left", padx=(8, 0))

        # ── Right: log panel ───────────────────────────────────────────────
        log_hdr = tk.Frame(right, bg=C["surface"])
        log_hdr.pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(log_hdr, text="OUTPUT LOG", bg=C["surface"],
                 fg=C["text_lo"], font=("Segoe UI", 7, "bold")).pack(side="left")

        self.text = tk.Text(right, bg=C["surface"], fg=C["text_med"],
                            font=("Consolas", 9), relief="flat", bd=0,
                            insertbackground=C["teal"], wrap="word",
                            selectbackground=C["overlay"],
                            selectforeground=C["text_hi"],
                            padx=10, pady=6)
        self.text.pack(fill="both", expand=True, padx=2, pady=(4, 2))
        self.text.configure(state="disabled")

        # colour tags for the log
        self.text.tag_configure("ok",    foreground=C["green"])
        self.text.tag_configure("error", foreground=C["red"])
        self.text.tag_configure("warn",  foreground=C["amber"])
        self.text.tag_configure("dim",   foreground=C["text_lo"])
        self.text.tag_configure("teal",  foreground=C["teal"])

        # ── Footer ─────────────────────────────────────────────────────────
        footer = tk.Frame(root_pad, bg=C["bg"])
        footer.pack(fill="x", pady=(10, 0))
        ps_path = resolve_powershell_path()
        tk.Label(footer, text=f"ps  {ps_path}", bg=C["bg"],
                 fg=C["text_lo"], font=("Consolas", 7)).pack(side="left")
        tk.Label(footer, text="Leon Priest · 7h3v01d · Apache 2.0",
                 bg=C["bg"], fg=C["text_lo"], font=("Segoe UI", 7)).pack(side="right")

    # ── State management ──────────────────────────────────────────────────────

    def _set_busy(self, busy: bool, msg: str = ""):
        if busy:
            self.status_var.set(msg or "Working…")
            self._status_dot.configure(fg=C["amber"])
            state = "disabled"
        else:
            self.status_var.set(msg or "Ready")
            self._status_dot.configure(fg=C["green"])
            state = "normal"
        for btn in (self.refresh_btn, self.run_btn, self.copy_btn):
            btn.configure(state=state)
        self.configure(cursor="watch" if busy else "")

    def _log(self, s: str, tag: str = ""):
        self.text.configure(state="normal")
        if tag:
            self.text.insert("end", s + "\n", tag)
        else:
            self.text.insert("end", s + "\n")
        self.text.see("end")
        self.text.configure(state="disabled")

    def _clear_log(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

    def _apply_scores(self, obj: dict):
        # Base score ring
        spr = obj.get("WinSPRLevel")
        self._ring.set_score(float(spr) if spr is not None else None)

        # Component bars
        for key, bar in self._score_bars.items():
            val = obj.get(key)
            bar.set_score(float(val) if val is not None else None)

        # Meta fields
        state_raw = obj.get("WinSATAssessmentState")
        self._meta_vars["WinSATAssessmentState"].set(
            decode_assessment_state(state_raw) if state_raw is not None else "—"
        )
        self._meta_vars["TimeTaken"].set(str(obj.get("TimeTaken", "—")))

    # ── Actions ───────────────────────────────────────────────────────────────

    def copy_json(self):
        if not self.raw_json:
            messagebox.showinfo(APP_TITLE, "No JSON yet — click Refresh first.")
            return
        self.clipboard_clear()
        self.clipboard_append(self.raw_json)
        self.status_var.set("JSON copied")

    def refresh_scores(self):
        def worker():
            self._clear_log()
            self._log("  querying Win32_WinSAT…", "dim")
            rc, out, err = run_powershell(ps_query_winsat_json(), timeout_sec=30)
            self.after(0, lambda: self._handle_query_result(rc, out, err))

        self._set_busy(True, "Querying…")
        threading.Thread(target=worker, daemon=True).start()

    def _handle_query_result(self, rc: int, out: str, err: str):
        try:
            if rc != 0:
                self._log(f"  [exit {rc}]", "error")
                if err: self._log(err, "error")
                if out: self._log(out, "warn")
                messagebox.showerror(APP_TITLE, "Failed to query Win32_WinSAT.\nSee log for details.")
                return

            if not out:
                self._log("  [empty output]", "error")
                messagebox.showerror(APP_TITLE, "PowerShell returned no output.")
                return

            self.raw_json = out
            self._log("  raw JSON →", "dim")
            self._log(out, "teal")

            obj = json.loads(out)
            if isinstance(obj, list) and obj:
                obj = obj[0]
            if not isinstance(obj, dict):
                raise ValueError("Unexpected JSON structure.")

            self._apply_scores(obj)
            self._log("\n  ✓ scores updated", "ok")
            self.status_var.set("Scores updated")
        except Exception as e:
            self._log(f"  [error] {e}", "error")
            messagebox.showerror(APP_TITLE, f"Parse error:\n{e}")
        finally:
            self._set_busy(False)

    def run_assessment(self):
        def worker():
            self._clear_log()
            self._log("  running winsat formal …", "warn")
            self._log("  (may require Administrator — this takes a minute)\n", "dim")
            rc, out, err = run_powershell(ps_run_winsat_assessment(), timeout_sec=600)

            def done():
                if rc != 0:
                    self._log(f"  [exit {rc}]", "error")
                    if err: self._log(err, "error")
                    if out: self._log(out, "warn")
                    messagebox.showerror(
                        APP_TITLE,
                        "WinSAT failed.\nTry running as Administrator.\nSee log for details.",
                    )
                    self._set_busy(False, "Ready")
                    return

                if out: self._log(out)
                if err: self._log(err, "warn")
                self._log("\n  re-querying…", "dim")
                self._set_busy(True, "Re-querying…")
                threading.Thread(target=self._refresh_after_assessment, daemon=True).start()

            self.after(0, done)

        self._set_busy(True, "Running assessment…")
        threading.Thread(target=worker, daemon=True).start()

    def _refresh_after_assessment(self):
        rc, out, err = run_powershell(ps_query_winsat_json(), timeout_sec=30)
        self.after(0, lambda: self._handle_query_result(rc, out, err))


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    app = WinSatGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
