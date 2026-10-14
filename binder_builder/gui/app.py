"""Main window: switches between the Start screen and a Binder screen (SPEC §12).

All binder logic lives in `engine.session.BinderSession`; this module and its
siblings only present it. Long work (plan, apply, update check) runs on a worker
thread and reports back through `after()`, so the window never freezes.
"""

from __future__ import annotations

import gc
import json
import os
import queue
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from .. import __version__
from ..engine.manifest import ManifestError
from ..engine.session import BinderSession

APP_NAME = "Binder Builder"


def app_data_dir() -> Path:
    base = Path(os.environ.get("APPDATA") or Path.home())
    d = base / "BinderBuilder"
    d.mkdir(parents=True, exist_ok=True)
    return d


class RecentBinders:
    MAX = 10

    def __init__(self, path: Path | None = None):
        self.path = path or app_data_dir() / "recent.json"

    def load(self) -> list[str]:
        try:
            items = json.loads(self.path.read_text(encoding="utf-8"))
            return [p for p in items if isinstance(p, str)]
        except (OSError, ValueError):
            return []

    def add(self, folder: Path) -> None:
        items = [str(folder)] + [p for p in self.load() if os.path.normcase(p) != os.path.normcase(str(folder))]
        self.path.write_text(json.dumps(items[: self.MAX], indent=2), encoding="utf-8")

    def remove(self, folder: str) -> None:
        self.path.write_text(json.dumps([p for p in self.load() if p != folder], indent=2), encoding="utf-8")


def _worker(fn, q: queue.Queue, tid: int) -> None:
    try:
        q.put((tid, fn(), None))
    except Exception as e:  # noqa: BLE001 - handed to the Tk thread
        q.put((tid, None, e))


def _check_updates():
    from .. import updater

    return updater.check_for_update(__version__)


class App(tk.Tk):
    def __init__(self, check_updates: bool = True):
        super().__init__()
        self.title(f"{APP_NAME} {__version__}")
        self.geometry("1200x760")
        self.minsize(900, 560)
        self._style()
        self.session: BinderSession | None = None
        self.view: ttk.Frame | None = None
        self.recent = RecentBinders()
        self.update_info = None
        self._tasks: queue.Queue = queue.Queue()
        self._callbacks: dict[int, tuple] = {}
        self._task_seq = 0
        self.report_callback_exception = self._report_error
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self._poll_tasks)
        self._gc_pending = False
        self.bind_all("<Destroy>", self._collect_soon, add="+")
        self.show_start()
        if check_updates:
            self.run_task(_check_updates, self._on_update_info, busy=None, on_error=lambda e: None)

    def _style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Heading.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Muted.TLabel", foreground="#666666")
        style.configure("Banner.TLabel", background="#FFF4CE", foreground="#5C4400", padding=8)
        style.configure("Banner.TFrame", background="#FFF4CE")
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", rowheight=24)

    # ---- background work ----------------------------------------------------------
    def run_task(self, fn, on_done, busy: str | None = "Working...", on_error=None) -> None:
        """Run `fn()` on a worker thread; call `on_done(result)` or `on_error(exc)` on the Tk thread.

        `fn` must not reference any Tk object: if the last reference to a widget were dropped on
        the worker thread, Tcl would crash. Callbacks stay here, keyed by a task id.
        """
        if busy:
            self.config(cursor="watch")
            self.set_status(busy)
        self._task_seq += 1
        tid = self._task_seq
        self._callbacks[tid] = (on_done, on_error, busy)
        threading.Thread(target=_worker, args=(fn, self._tasks, tid), daemon=True).start()

    def _poll_tasks(self) -> None:
        try:
            while True:
                tid, result, err = self._tasks.get_nowait()
                on_done, on_error, busy = self._callbacks.pop(tid, (None, None, None))
                if busy:
                    self.config(cursor="")
                    self.set_status("")
                if err is not None:
                    if on_error:
                        on_error(err)
                    else:
                        self._report_error(type(err), err, err.__traceback__)
                elif on_done:
                    on_done(result)
        except queue.Empty:
            pass
        self.after(100, self._poll_tasks)

    def _collect_soon(self, _event=None) -> None:
        """Destroyed screens leave Tk Variables in reference cycles. Collect them here, on the Tk
        thread, before a worker thread's allocation can trigger a GC that frees them there."""
        if not self._gc_pending:
            self._gc_pending = True

            def run():
                self._gc_pending = False
                gc.collect()

            self.after_idle(run)

    def set_status(self, text: str) -> None:
        if self.view is not None and hasattr(self.view, "set_busy"):
            self.view.set_busy(text)

    # ---- errors -----------------------------------------------------------------
    def _report_error(self, exc_type, exc, tb) -> None:
        log = app_data_dir() / "error.log"
        try:
            with log.open("a", encoding="utf-8") as f:
                f.write("".join(traceback.format_exception(exc_type, exc, tb)) + "\n")
        except OSError:
            pass
        if isinstance(exc, (ValueError, ManifestError, PermissionError, FileNotFoundError)):
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
        else:
            messagebox.showerror(APP_NAME, f"Something went wrong:\n\n{exc}\n\nDetails were saved to {log}",
                                 parent=self)

    # ---- screens ------------------------------------------------------------------
    def _swap(self, view: ttk.Frame) -> None:
        if self.view is not None:
            self.view.destroy()
            gc.collect()
        self.view = view
        view.pack(fill="both", expand=True)

    def show_start(self) -> None:
        from .start import StartView

        self.title(f"{APP_NAME} {__version__}")
        self._swap(StartView(self))

    def open_binder(self, folder: str | Path) -> None:
        folder = Path(folder)
        if not folder.is_dir():
            messagebox.showwarning(APP_NAME, f"The folder {folder} no longer exists.", parent=self)
            self.recent.remove(str(folder))
            self.show_start()
            return
        if not BinderSession.is_binder(folder):
            if not messagebox.askyesno(
                APP_NAME,
                f"{folder.name} is not a Binder Builder binder yet.\n\n"
                "Make it a new, empty binder? Nothing in the folder is changed; files already there "
                "will be listed as orphans until you add them to the Table of Contents.",
                parent=self,
            ):
                return
            BinderSession.create(folder, folder.name).close()
        self.close_binder(show_start=False)
        self.session = BinderSession(folder)
        self.recent.add(folder)
        from .binder_view import BinderView

        self.title(f"{self.session.binder.name} - {APP_NAME}")
        self._swap(BinderView(self, self.session))

    def choose_and_open(self) -> None:
        folder = filedialog.askdirectory(parent=self, title="Open a binder folder", mustexist=True)
        if folder:
            self.open_binder(folder)

    def new_binder(self) -> None:
        parent = filedialog.askdirectory(parent=self, title="Where should the new binder folder go?", mustexist=True)
        if not parent:
            return
        name = simpledialog.askstring(APP_NAME, "Binder name:", parent=self)
        if not name or not name.strip():
            return
        folder = Path(parent) / name.strip()
        if folder.exists() and any(folder.iterdir()):
            if not messagebox.askyesno(APP_NAME, f"{folder} already exists and is not empty. Use it anyway?",
                                       parent=self):
                return
        BinderSession.create(folder, name.strip()).close()
        self.open_binder(folder)

    def close_binder(self, show_start: bool = True) -> None:
        if self.session is not None:
            self.session.close()
            self.session = None
        if show_start:
            self.show_start()

    def on_close(self) -> None:
        self.close_binder(show_start=False)
        self.destroy()

    # ---- updates ------------------------------------------------------------------
    def _on_update_info(self, info) -> None:
        self.update_info = info
        if info and self.view is not None and hasattr(self.view, "show_update"):
            self.view.show_update(info)

    def offer_update(self) -> None:
        from .. import updater

        info = self.update_info
        if not info:
            return
        notes = info.notes.strip() or "(no release notes)"
        if not updater.can_self_update():
            messagebox.showinfo(APP_NAME, f"Version {info.version} is available.\n\n{notes}\n\n"
                                          f"Download it from:\n{info.page_url}", parent=self)
            return
        if messagebox.askyesno(APP_NAME, f"Update to version {info.version}?\n\n{notes}\n\n"
                                         "The app will close, update itself and reopen.", parent=self):
            self.close_binder(show_start=False)
            self.run_task(lambda: updater.download_and_stage(info), self._finish_update, busy="Downloading update...")

    def _finish_update(self, staged) -> None:
        from .. import updater

        updater.launch_swap_and_exit(staged)
        self.destroy()


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
