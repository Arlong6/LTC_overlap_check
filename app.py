import io
import os
import re
import sys
import json
import queue
import threading
import subprocess

import customtkinter as ctk
from tkinter import filedialog, messagebox

from main import LTC_solution

import sys as _sys
APP_DIR = (os.path.dirname(_sys.executable) if getattr(_sys, "frozen", False)
           else os.path.dirname(os.path.abspath(__file__)))

# ── Config ───────────────────────────────────────────────────────────────────

CONFIG_FILE = os.path.join(APP_DIR, "ltc_config.json")


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ── Text redirector (thread-safe via queue) ──────────────────────────────────

class TextRedirector(io.TextIOBase):
    """Redirects writes to an app queue so the main thread updates the textbox."""
    _ANSI_ESCAPE = re.compile(r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])")
    _RICH_MARKUP = re.compile(r"\[/?[a-zA-Z0-9_ /]+\]")

    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, text: str):
        clean = self._ANSI_ESCAPE.sub("", text)
        clean = self._RICH_MARKUP.sub("", clean)
        if clean:
            self._q.put(("log", clean))
        return len(text)

    def flush(self):
        pass


# ── Main Application ─────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("LTC 重疊檢查工具")
        self.geometry("1150x720")
        self.minsize(900, 600)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self._config = load_config()
        self._q: queue.Queue = queue.Queue()
        self._running = False
        self._orig_stdout = None
        self._results_row = 0

        self._build_ui()
        self._load_paths()
        self._poll_queue()

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=0, minsize=390)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_left()
        self._build_right()

    def _build_left(self):
        left = ctk.CTkFrame(self, width=390)
        left.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        left.grid_propagate(False)
        left.grid_columnconfigure(0, weight=1)

        r = 0

        # Title
        ctk.CTkLabel(
            left, text="LTC 重疊檢查", font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=r, column=0, pady=(18, 14), padx=16, sticky="w")
        r += 1

        # Folder pickers
        self._path_vars: dict[str, ctk.StringVar] = {}
        # required=True → must be filled; required=False → optional
        picker_defs = [
            ("csv_path",     "服務紀錄資料夾", True),
            ("support_path", "支援資料資料夾", False),
            ("class_path",   "上課資料資料夾", False),
            ("save_path",    "結果儲存位置",   True),
        ]
        for key, label, required in picker_defs:
            grp = ctk.CTkFrame(left, fg_color="transparent")
            grp.grid(row=r, column=0, sticky="ew", padx=12, pady=3)
            grp.grid_columnconfigure(1, weight=1)

            display = f"{label} *" if required else f"{label}（選填）"
            ctk.CTkLabel(grp, text=display, width=130, anchor="w").grid(
                row=0, column=0, padx=(0, 6)
            )
            var = ctk.StringVar()
            self._path_vars[key] = var
            ctk.CTkEntry(grp, textvariable=var).grid(row=0, column=1, sticky="ew")
            ctk.CTkButton(
                grp, text="選擇", width=52,
                command=lambda k=key: self._browse_folder(k),
            ).grid(row=0, column=2, padx=(6, 0))
            r += 1

        # Divider
        ctk.CTkFrame(left, height=2, fg_color=("gray70", "gray40")).grid(
            row=r, column=0, sticky="ew", padx=12, pady=12
        )
        r += 1

        # Run button
        self._run_btn = ctk.CTkButton(
            left, text="▶  開始執行", height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._start_run,
        )
        self._run_btn.grid(row=r, column=0, padx=12, sticky="ew")
        r += 1

        # Progress label
        self._progress_label = ctk.CTkLabel(left, text="", anchor="w")
        self._progress_label.grid(row=r, column=0, padx=14, pady=(10, 2), sticky="ew")
        r += 1

        # Progress bar
        self._progress_bar = ctk.CTkProgressBar(left)
        self._progress_bar.set(0)
        self._progress_bar.grid(row=r, column=0, padx=12, sticky="ew")
        r += 1

        # Log label
        ctk.CTkLabel(left, text="執行紀錄", anchor="w").grid(
            row=r, column=0, padx=14, pady=(14, 2), sticky="w"
        )
        r += 1

        # Log textbox (expandable)
        self._log = ctk.CTkTextbox(left, state="disabled", wrap="word", font=ctk.CTkFont(size=12))
        self._log.grid(row=r, column=0, padx=12, pady=(0, 12), sticky="nsew")
        left.grid_rowconfigure(r, weight=1)

    def _build_right(self):
        right = ctk.CTkFrame(self)
        right.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            right, text="檢查結果", font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=(18, 8), padx=16, sticky="w")

        self._results_frame = ctk.CTkScrollableFrame(right, corner_radius=0)
        self._results_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        self._results_frame.grid_columnconfigure(0, weight=1)
        # Clamp elastic scroll on macOS so content doesn't bounce past edges
        canvas = self._results_frame._parent_canvas
        canvas.configure(yscrollincrement=1)
        canvas.bind("<MouseWheel>", self._clamp_scroll, add="+")

        self._open_folder_btn = ctk.CTkButton(
            right, text="📁  開啟存檔資料夾", height=38,
            command=self._open_save_folder,
            state="disabled",
        )
        self._open_folder_btn.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="ew")

    # ── Path helpers ─────────────────────────────────────────────────────

    def _browse_folder(self, key: str):
        current = self._path_vars[key].get().strip()
        initial = current if current and os.path.isdir(current) else APP_DIR
        folder = filedialog.askdirectory(initialdir=initial)
        if folder:
            self._path_vars[key].set(folder)
            self._config[key] = folder
            save_config(self._config)

    def _load_paths(self):
        for key, var in self._path_vars.items():
            if key in self._config:
                var.set(self._config[key])

    # ── Run logic ─────────────────────────────────────────────────────────

    def _start_run(self):
        if self._running:
            return

        paths = {k: v.get().strip() for k, v in self._path_vars.items()}
        missing = [
            label
            for key, label in [("csv_path", "服務紀錄"), ("save_path", "儲存位置")]
            if not paths[key]
        ]
        if missing:
            messagebox.showwarning("缺少路徑", f"請先選擇以下路徑：\n{chr(10).join(missing)}")
            return

        # Ensure save path exists
        os.makedirs(paths["save_path"], exist_ok=True)

        # Save config
        for k, v in paths.items():
            self._config[k] = v
        save_config(self._config)

        # Clear results panel
        for w in self._results_frame.winfo_children():
            w.destroy()
        self._results_row = 0
        self._open_folder_btn.configure(state="disabled")

        # Reset log & progress
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        self._progress_bar.set(0)
        self._progress_label.configure(text="準備中…")

        self._running = True
        self._run_btn.configure(state="disabled", text="執行中…")

        # Redirect stdout to queue-based redirector
        self._orig_stdout = sys.stdout
        sys.stdout = TextRedirector(self._q)

        ltc = LTC_solution(
            paths["csv_path"], paths["save_path"],
            paths["support_path"], paths["class_path"],
        )
        threading.Thread(target=self._worker, args=(ltc,), daemon=True).start()

    def _worker(self, ltc: LTC_solution):
        try:
            ltc.run(
                on_progress=lambda *a: self._q.put(("progress", a)),
                on_result=lambda *a: self._q.put(("result", a)),
            )
            self._q.put(("done", None))
        except Exception as exc:
            import traceback
            self._q.put(("error", traceback.format_exc()))

    # ── Queue polling (main thread) ───────────────────────────────────────

    def _poll_queue(self):
        try:
            while True:
                kind, data = self._q.get_nowait()
                if kind == "log":
                    self._append_log(data)
                elif kind == "progress":
                    self._handle_progress(*data)
                elif kind == "result":
                    self._handle_result(*data)
                elif kind == "done":
                    self._on_done()
                elif kind == "error":
                    self._on_error(data)
        except queue.Empty:
            pass
        finally:
            self.after(80, self._poll_queue)

    def _append_log(self, text: str):
        self._log.configure(state="normal")
        self._log.insert("end", text)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _handle_progress(self, check_type: str, current: int, total: int, name: str):
        pct = current / total if total > 0 else 0
        self._progress_bar.set(pct)
        self._progress_label.configure(text=f"{check_type}：{current}/{total}  {name}")

    def _handle_result(self, check_type: str, display_name: str, count: int, path: str):
        has_issue = count > 0
        icon = "⚠" if has_issue else "✓"
        fg_color = ("#FDECEA", "#4A1E1E") if has_issue else ("#E6F4EA", "#1A3A22")
        icon_color = "#C0392B" if has_issue else "#27AE60"

        type_labels = {"patient": "個案", "worker": "居服員", "support": "支援", "class": "上課"}
        type_text = type_labels.get(check_type, check_type)

        card = ctk.CTkFrame(self._results_frame, fg_color=fg_color)
        card.grid(row=self._results_row, column=0, sticky="ew", padx=4, pady=3)
        card.grid_columnconfigure(2, weight=1)
        self._results_row += 1

        # Icon + type label
        ctk.CTkLabel(
            card, text=f" {icon} ", text_color=icon_color,
            font=ctk.CTkFont(size=16, weight="bold"), width=36,
        ).grid(row=0, column=0, padx=(8, 0), pady=10)

        ctk.CTkLabel(
            card, text=type_text, font=ctk.CTkFont(weight="bold"),
            width=64, anchor="w",
        ).grid(row=0, column=1, padx=(2, 8), pady=10)

        # Summary
        count_text = f"{count} 筆有問題" if has_issue else "無問題"
        summary = f"{display_name}　{count_text}"
        ctk.CTkLabel(card, text=summary, anchor="w").grid(
            row=0, column=2, padx=4, pady=10, sticky="ew"
        )

        # Open report button
        if path and os.path.exists(path):
            def _open_report(p=path):
                self._append_log(f"開啟：{p}\n")
                self._open_path(p)
            ctk.CTkButton(
                card, text="開報告", width=72,
                command=_open_report,
            ).grid(row=0, column=3, padx=(4, 10), pady=8)

        self._open_folder_btn.configure(state="normal")

    def _on_done(self):
        self._restore_stdout()
        self._running = False
        self._run_btn.configure(state="normal", text="▶  開始執行")
        self._progress_bar.set(1)
        self._progress_label.configure(text="✓ 執行完成")
        self._open_folder_btn.configure(state="normal")

    def _on_error(self, msg: str):
        self._restore_stdout()
        self._running = False
        self._run_btn.configure(state="normal", text="▶  開始執行")
        self._progress_label.configure(text="✗ 執行失敗")
        self._append_log(f"\n[執行錯誤]\n{msg}\n")

    def _restore_stdout(self):
        if self._orig_stdout is not None:
            sys.stdout = self._orig_stdout
            self._orig_stdout = None

    def _clamp_scroll(self, _event=None):
        """Prevent elastic over-scroll past top/bottom edges."""
        canvas = self._results_frame._parent_canvas
        lo, hi = canvas.yview()
        if lo < 0.0:
            canvas.yview_moveto(0.0)
        elif hi > 1.0:
            canvas.yview_moveto(1.0 - (hi - lo))

    # ── OS helpers ────────────────────────────────────────────────────────

    def _open_path(self, path: str):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["/usr/bin/open", path], check=True)
            else:
                subprocess.run(["xdg-open", path], check=True)
        except Exception as e:
            messagebox.showerror("開啟失敗", f"無法開啟檔案：\n{path}\n\n{e}")

    def _open_save_folder(self):
        save_path = self._path_vars["save_path"].get().strip()
        if save_path and os.path.exists(save_path):
            self._open_path(save_path)
        else:
            messagebox.showwarning("找不到路徑", "存檔資料夾不存在，請先執行一次檢查。")


if __name__ == "__main__":
    app = App()
    app.mainloop()
