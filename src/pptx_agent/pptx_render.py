"""Renders a .pptx to one PNG per slide via PowerPoint COM automation, so the
Visual QA Agent can look at the actual deliverable instead of the HTML
intermediate. (The spec's own architecture calls for a "render -> QA ->
fix" loop via LibreOffice headless; this machine has PowerPoint but not
LibreOffice, and rendering through the real target application is at least
as faithful a QA signal.) Windows + PowerPoint only.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import time
from pathlib import Path

PP_SAVE_AS_PNG = 18


def _list_powerpnt_pids() -> set[int]:
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
            text=True,
            stderr=subprocess.DEVNULL,
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        return set()
    pids = set()
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[0].upper().startswith("POWERPNT"):
            try:
                pids.add(int(parts[1]))
            except ValueError:
                pass
    return pids


def _force_kill(pid: int) -> None:
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _render_once(pptx_path: Path, out_dir: Path) -> None:
    import win32com.client

    # DispatchEx starts a fresh, isolated PowerPoint process rather than
    # attaching to whatever instance the user (or a previous CLI run's
    # --open-result) already has open. We also track its PID so we can force
    # -kill exactly that process afterward: app.Quit() has been observed to
    # silently fail here, leaving an orphaned instance that blocks the next
    # Presentations.Open() with an opaque COM error. Never touches PIDs that
    # existed before this call (e.g. the user's own PowerPoint window).
    pids_before = _list_powerpnt_pids()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(pptx_path.resolve()), WithWindow=False)
        try:
            pres.SaveAs(str(out_dir.resolve()), PP_SAVE_AS_PNG)
        finally:
            pres.Close()
    finally:
        try:
            app.Quit()
        except Exception:
            pass  # best-effort; the PID-based cleanup below is the real guarantee
        time.sleep(0.5)
        for pid in _list_powerpnt_pids() - pids_before:
            _force_kill(pid)


def render_pptx_to_pngs(pptx_path: Path, out_dir: Path, *, retries: int = 3) -> list[Path]:
    """PowerPoint COM automation is prone to transient failures right after a
    prior instance was torn down (DCOM/Office cleanup timing). Retry a few
    times with a short backoff before giving up."""
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            _render_once(pptx_path, out_dir)
            last_error = None
            break
        except Exception as exc:
            last_error = exc
            time.sleep(2 * attempt)
    if last_error is not None:
        raise last_error

    def slide_number(p: Path) -> int:
        m = re.search(r"(\d+)", p.stem)
        return int(m.group(1)) if m else 0

    # Windows filesystems are case-insensitive, so globbing both "*.PNG" and
    # "*.png" would double-count every file.
    pngs = list(out_dir.glob("*.png"))
    pngs.sort(key=slide_number)
    return pngs


def powerpoint_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    try:
        import win32com.client

        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Quit()
        return True
    except Exception:
        return False
