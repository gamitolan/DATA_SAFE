"""
app.py
Single Tkinter application for the OFS-based M580 tag backup/restore
tool. Choose Backup or Restore, pick the relevant Excel file(s), Run.

Backup discovers its tag list directly from OFS (via its address space
browser) - there is no tag list file to prepare or maintain.

The actual backup/restore work (backup.run_backup / restore.run_restore)
runs on a background thread, so the window stays responsive - and
shows live progress - during long binding/reading/writing passes.
Tkinter itself is not thread-safe, so the worker thread never touches
widgets directly: it posts messages onto a thread-safe queue, and the
main thread drains that queue on a timer via root.after(). The one
tricky part is the restore confirmation dialog, which must be shown on
the main thread but needs to block the worker thread until answered -
handled with a threading.Event handshake (see _poll_queue / gui_confirm).

This is the file to run:  python app.py
"""

import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from config import load_config
from ofs_client import OFSClient
import backup
import restore


class App:
    def __init__(self, root):
        self.root = root
        root.title("M580 Tag Backup / Restore (via OFS)")
        root.geometry("640x640")
        root.minsize(600, 560)

        self.mode = tk.StringVar(value="backup")
        self.backup_out_path = tk.StringVar()
        self.restore_backup_path = tk.StringVar()
        self.restore_report_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Idle.")
        self.include_fb_members = tk.BooleanVar(value=False)
        self.device_alias = tk.StringVar(value="(All devices)")

        # State for the "..." animation shown during dead time before a
        # phase has anything measurable to report yet (e.g. discovery's
        # connection/startup time, before the browse finds its first item).
        self._animating = False
        self._animation_phase = ""
        self._dots_count = 0

        cfg = load_config()
        self.ofs_node = tk.StringVar(value=cfg.get("ofs_node", "localhost"))
        self.ofs_progid = tk.StringVar(value=cfg.get("ofs_progid", "Schneider-Aut.OFS"))
        self._godsharp_netapi_lib_dir = cfg.get("godsharp_netapi_lib_dir")
        self._godsharp_base_lib_dir = cfg.get("godsharp_base_lib_dir")
        self._titanium_lib_dir = cfg.get("titanium_lib_dir")

        # Queue of messages FROM the worker thread TO the main thread.
        # Only the main thread ever touches Tkinter widgets.
        self._queue = queue.Queue()
        self._worker_thread = None

        # Handshake for the restore confirmation dialog: the worker
        # thread posts a "confirm" message and then blocks on this
        # Event; the main thread shows the dialog (safe, it's the
        # main thread), stores the answer, and sets the Event to
        # release the worker.
        self._confirm_event = threading.Event()
        self._confirm_result = False

        self._build_ui()
        self._on_mode_change()
        self._poll_queue()
        self._on_refresh_devices()  # populate the device list right away, no manual click needed

    # ---------- UI construction ----------

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        mode_frame = ttk.LabelFrame(self.root, text="Mode")
        mode_frame.pack(fill="x", **pad)
        ttk.Radiobutton(mode_frame, text="Backup (discover tags from OFS, read from CPU -> Excel)",
                         variable=self.mode, value="backup",
                         command=self._on_mode_change).pack(side="left", padx=8, pady=6)
        ttk.Radiobutton(mode_frame, text="Restore (Excel -> write to CPU)",
                         variable=self.mode, value="restore",
                         command=self._on_mode_change).pack(side="left", padx=8, pady=6)

        # Backup-specific fields
        self.backup_frame = ttk.LabelFrame(self.root, text="Backup output")
        self._file_row(self.backup_frame, "Backup output file (.xlsx):",
                        self.backup_out_path, self._browse_backup_out)

        ttk.Checkbutton(self.backup_frame,
                         text="Include function block instance data (slower; some fields will "
                              "predictably fail to read/write - see log)",
                         variable=self.include_fb_members).pack(anchor="w", padx=8, pady=(0, 4))
        ttk.Label(self.backup_frame,
                  text="Tags are discovered directly from OFS - no input file needed.",
                  foreground="gray").pack(anchor="w", padx=8, pady=(0, 6))

        # Restore-specific fields
        self.restore_frame = ttk.LabelFrame(self.root, text="Restore files")
        self._file_row(self.restore_frame, "Backup file to restore from (.xlsx):",
                        self.restore_backup_path, self._browse_restore_backup)
        self._file_row(self.restore_frame, "Restore report output (.xlsx, optional):",
                        self.restore_report_path, self._browse_restore_report)

        # Connection settings
        conn_frame = ttk.LabelFrame(self.root, text="OFS connection settings")
        conn_frame.pack(fill="x", **pad)
        self._entry_row(conn_frame, "OFS node:", self.ofs_node)
        self._entry_row(conn_frame, "OFS ProgID:", self.ofs_progid)

        device_row = ttk.Frame(conn_frame)
        device_row.pack(fill="x", padx=8, pady=4)
        ttk.Label(device_row, text="Device:", width=24, anchor="w").pack(side="left")
        self.device_combo = ttk.Combobox(device_row, textvariable=self.device_alias,
                                          values=["(All devices)"], state="readonly")
        self.device_combo.pack(side="left", fill="x", expand=True, padx=4)
        self.refresh_devices_button = ttk.Button(device_row, text="Refresh",
                                                  command=self._on_refresh_devices)
        self.refresh_devices_button.pack(side="left")

        # Run button
        self.run_button = ttk.Button(self.root, text="Run", command=self._on_run)
        self.run_button.pack(pady=(8, 2))

        # Progress bar + status line - shown/updated during a run so
        # the window never looks idle/frozen during a long pass.
        progress_frame = ttk.Frame(self.root)
        progress_frame.pack(fill="x", padx=8, pady=(0, 6))
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress_bar.pack(fill="x")
        ttk.Label(progress_frame, textvariable=self.status_text, foreground="gray").pack(anchor="w")

        # Log area
        log_frame = ttk.LabelFrame(self.root, text="Log")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log_widget = scrolledtext.ScrolledText(log_frame, height=14, state="disabled", wrap="word")
        self.log_widget.pack(fill="both", expand=True, padx=4, pady=4)

    def _file_row(self, parent, label, var, browse_cmd):
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text=label, width=32, anchor="w").pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row, text="Browse...", command=browse_cmd).pack(side="left")

    def _entry_row(self, parent, label, var):
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text=label, width=24, anchor="w").pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=4)

    def _on_mode_change(self):
        pad = {"padx": 8, "pady": 6}
        if self.mode.get() == "backup":
            self.restore_frame.pack_forget()
            self.backup_frame.pack(fill="x", **pad)
        else:
            self.backup_frame.pack_forget()
            self.restore_frame.pack(fill="x", **pad)

    # ---------- File pickers ----------

    def _browse_backup_out(self):
        path = filedialog.asksaveasfilename(title="Save backup file as",
                                             defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.backup_out_path.set(path)

    def _browse_restore_backup(self):
        path = filedialog.askopenfilename(title="Select backup file to restore from",
                                           filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.restore_backup_path.set(path)

    def _browse_restore_report(self):
        path = filedialog.asksaveasfilename(title="Save restore report as",
                                             defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.restore_report_path.set(path)

    # ---------- Logging (called from the MAIN thread only) ----------

    def log(self, message):
        self.log_widget.configure(state="normal")
        self.log_widget.insert("end", str(message) + "\n")
        self.log_widget.see("end")
        self.log_widget.configure(state="disabled")

    # ---------- Queue-based communication with the worker thread ----------
    #
    # The worker thread never touches Tkinter widgets directly. It
    # calls these three small callables, which just post a message
    # onto the queue (thread-safe) and, for confirm, block until the
    # main thread answers.

    def _worker_log(self, message):
        self._queue.put(("log", message))

    def _worker_progress(self, phase, current, total):
        self._queue.put(("progress", phase, current, total))

    def _worker_confirm(self, tag_value_pairs):
        self._confirm_event.clear()
        self._queue.put(("confirm", tag_value_pairs))
        self._confirm_event.wait()  # blocks the WORKER thread, not the GUI
        return self._confirm_result

    def _poll_queue(self):
        """Runs on the main thread via root.after() - drains whatever
        the worker thread has posted since the last tick and applies
        it to the UI. Reschedules itself, so it runs continuously for
        the life of the app, not just during a run."""
        try:
            while True:
                item = self._queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self.log(item[1])
                elif kind == "progress":
                    _, phase, current, total = item
                    if total:
                        self._animating = False
                        pct = int(100 * current / total)
                        self.progress_bar["value"] = pct
                        self.status_text.set(f"{phase}: {current}/{total}")
                    elif current:
                        # Real progress with an unknown total (e.g. discovery
                        # mid-browse) - stop any dead-time animation and show it.
                        self._animating = False
                        self.progress_bar["value"] = 0
                        self.status_text.set(f"{phase}... ({current})")
                    else:
                        # current == 0, total == 0: nothing measurable yet
                        # (e.g. discovery hasn't found its first item) - show
                        # a moving "..." instead of a frozen, silent status.
                        self.progress_bar["value"] = 0
                        self._animation_phase = phase
                        if not self._animating:
                            self._animating = True
                            self._dots_count = 0
                            self._animate_dots()
                elif kind == "devices":
                    _, aliases = item
                    self._animating = False
                    values = ["(All devices)"] + list(aliases)
                    self.device_combo["values"] = values
                    if self.device_alias.get() not in values:
                        self.device_alias.set("(All devices)")
                    self.log(f"Found {len(aliases)} device alias(es): "
                              f"{', '.join(aliases) if aliases else '(none)'}")
                    self.status_text.set("Idle.")
                    self.run_button.configure(state="normal")
                    self.refresh_devices_button.configure(state="normal")
                elif kind == "confirm":
                    _, tag_value_pairs = item
                    self._confirm_result = self._show_confirm_dialog(tag_value_pairs)
                    self._confirm_event.set()
                elif kind == "done":
                    self._animating = False
                    self.status_text.set("Done.")
                    self.progress_bar["value"] = 0
                    self.run_button.configure(state="normal")
                elif kind == "error":
                    self._animating = False
                    self.status_text.set("Error.")
                    self.progress_bar["value"] = 0
                    self.log(f"ERROR: {item[1]}")
                    messagebox.showerror("Error", item[1])
                    self.run_button.configure(state="normal")
                    self.refresh_devices_button.configure(state="normal")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _animate_dots(self):
        """Cycles '.', '..', '...' in the status line while a phase has
        nothing measurable to report yet. Stops rescheduling itself as
        soon as self._animating goes False (set the moment real progress,
        a done, or an error message arrives)."""
        if not self._animating:
            return
        dots = "." * ((self._dots_count % 3) + 1)
        self.status_text.set(f"{self._animation_phase}{dots}")
        self._dots_count += 1
        self.root.after(400, self._animate_dots)

    def _show_confirm_dialog(self, tag_value_pairs):
        message = (
            f"{len(tag_value_pairs)} tag(s) will be restored and written back to PLC memory.\n\n"
            f"Confirm?"
        )
        return messagebox.askyesno("Confirm restore write", message)

    # ---------- Run ----------

    def _current_config(self):
        return {
            "ofs_progid": self.ofs_progid.get().strip(),
            "ofs_node": self.ofs_node.get().strip(),
            "group_name": "BackupRestoreGroup",
            "godsharp_netapi_lib_dir": self._godsharp_netapi_lib_dir,
            "godsharp_base_lib_dir": self._godsharp_base_lib_dir,
            "titanium_lib_dir": self._titanium_lib_dir,
        }

    def _on_run(self):
        if self.mode.get() == "backup":
            if not self.backup_out_path.get():
                messagebox.showerror("Missing file", "Select where to save the backup file.")
                return
        else:
            if not self.restore_backup_path.get():
                messagebox.showerror("Missing file", "Select a backup file to restore from.")
                return

        config = self._current_config()
        self.run_button.configure(state="disabled")
        self.progress_bar["value"] = 0
        self.status_text.set("Starting...")

        if self.mode.get() == "backup":
            self._worker_log("=== Starting backup ===")
            self._worker_thread = threading.Thread(
                target=self._backup_worker, args=(config,), daemon=True)
        else:
            self._worker_log("=== Starting restore ===")
            self._worker_thread = threading.Thread(
                target=self._restore_worker, args=(config,), daemon=True)
        self._worker_thread.start()

    # ---------- Worker thread bodies (run OFF the main thread) ----------
    #
    # These call straight into backup.run_backup / restore.run_restore
    # with callables that only ever post to the queue (_worker_log /
    # _worker_progress / _worker_confirm) - never touching a Tkinter
    # widget directly, since that would not be thread-safe.

    def _backup_worker(self, config):
        try:
            device_alias = self.device_alias.get()
            if device_alias == "(All devices)":
                device_alias = None
            backup.run_backup(config, self.backup_out_path.get(),
                               log=self._worker_log, progress=self._worker_progress,
                               include_fb_members=self.include_fb_members.get(),
                               device_alias=device_alias)
            self._queue.put(("done",))
        except Exception as exc:  # noqa: BLE001 - surface any failure to the operator
            self._queue.put(("error", str(exc)))

    def _on_refresh_devices(self):
        config = self._current_config()
        self.run_button.configure(state="disabled")
        self.refresh_devices_button.configure(state="disabled")
        self._worker_log("Refreshing device list...")
        t = threading.Thread(target=self._refresh_devices_worker, args=(config,), daemon=True)
        t.start()

    def _refresh_devices_worker(self, config):
        try:
            client = OFSClient(node=config["ofs_node"], progid=config["ofs_progid"],
                                group_name=config["group_name"],
                                netapi_lib_dir=config["godsharp_netapi_lib_dir"],
                                base_lib_dir=config["godsharp_base_lib_dir"],
                                titanium_lib_dir=config["titanium_lib_dir"])
            aliases = client.list_device_aliases()
            self._queue.put(("devices", aliases))
        except Exception as exc:  # noqa: BLE001
            self._queue.put(("error", str(exc)))

    def _restore_worker(self, config):
        try:
            report_path = self.restore_report_path.get() or None
            restore.run_restore(config, self.restore_backup_path.get(), report_path,
                                 log=self._worker_log, confirm=self._worker_confirm,
                                 progress=self._worker_progress)
            self._queue.put(("done",))
        except Exception as exc:  # noqa: BLE001 - surface any failure to the operator
            self._queue.put(("error", str(exc)))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
