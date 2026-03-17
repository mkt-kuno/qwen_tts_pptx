from __future__ import annotations

import logging
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from typing import Any, cast

from app.common.paths import WorkspacePaths


LOG_FORMAT = "%(asctime)s %(levelname)s %(message)s"
logger = logging.getLogger(__name__)


class QueueLogHandler(logging.Handler):
    def __init__(self, event_queue: queue.Queue[tuple[str, str]]) -> None:
        super().__init__()
        self.event_queue = event_queue

    def emit(self, record: logging.LogRecord) -> None:
        try:
            message = self.format(record)
        except Exception:
            message = record.getMessage()
        self.event_queue.put(("log", message))


class MovieBuilderGui:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.project_root = Path.cwd().resolve()
        self.paths = WorkspacePaths.from_root(self.project_root)
        self.paths.ensure_directories()
        self.last_open_dir = self.paths.output
        self.event_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.root.title("PowerPoint TTS Movie Builder")
        self.root.geometry("980x760")
        self.root.minsize(900, 700)

        guessed_pptx = self._guess_pptx()
        guessed_ref_audio = self._guess_voice_audio()
        guessed_ref_text = self._guess_voice_text(guessed_ref_audio)
        self.guessed_pptx = guessed_pptx
        self.guessed_ref_audio = guessed_ref_audio
        self.guessed_ref_text = guessed_ref_text

        self.pptx_var = tk.StringVar(value=guessed_pptx)
        self.ref_audio_var = tk.StringVar(value=guessed_ref_audio)
        self.ref_text_var = tk.StringVar(value=guessed_ref_text)
        self.model_size_var = tk.StringVar(value="1.7B")
        self.device_var = tk.StringVar()
        self.dtype_var = tk.StringVar(value="auto")
        self.slide_width_var = tk.StringVar(value="2560")
        self.slide_height_var = tk.StringVar(value="1440")
        self.slide_padding_var = tk.StringVar(value="1.5")
        self.fps_var = tk.StringVar(value="5")
        self.language_vars = {
            "JA": tk.BooleanVar(value=True),
            "EN": tk.BooleanVar(value=True),
            "ZH": tk.BooleanVar(value=True),
        }
        self.status_var = tk.StringVar(value="Ready")

        self._configure_logging()
        self._build_layout()
        self._poll_events()

    def _configure_logging(self) -> None:
        self.log_handler = QueueLogHandler(self.event_queue)
        self.log_handler.setFormatter(logging.Formatter(LOG_FORMAT))
        self.stdout_handler = logging.StreamHandler(sys.stdout)
        self.stdout_handler.setFormatter(logging.Formatter(LOG_FORMAT))
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(self.log_handler)
        root_logger.addHandler(self.stdout_handler)

    def _build_layout(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        container = ttk.Frame(self.root, padding=12)
        container.grid(sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)

        self._build_inputs_frame(container)
        self._build_options_frame(container)
        self._build_actions_frame(container)
        self._build_status_frame(container)
        self._build_log_frame(container)

    def _build_inputs_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Inputs", padding=12)
        frame.grid(row=0, column=0, sticky="ew")
        frame.columnconfigure(1, weight=1)

        self._add_file_row(
            frame,
            row=0,
            label="PPTX",
            variable=self.pptx_var,
            filetypes=[("PowerPoint", "*.pptx")],
        )
        self._add_file_row(
            frame,
            row=1,
            label="Ref Audio",
            variable=self.ref_audio_var,
            filetypes=[("Audio", "*.wav *.mp3 *.flac *.ogg *.m4a"), ("All", "*.*")],
        )
        self._add_file_row(
            frame,
            row=2,
            label="Ref Text",
            variable=self.ref_text_var,
            filetypes=[("Text", "*.txt"), ("All", "*.*")],
        )

        hint = (
            "Ref Audio and Ref Text are required. The GUI uses the selected file paths "
            "directly and does not read or write .pt prompt cache files."
        )
        ttk.Label(frame, text=hint, foreground="#555555").grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )

    def _build_options_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Options", padding=12)
        frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        for column in range(4):
            frame.columnconfigure(column, weight=1)

        ttk.Label(frame, text="Languages").grid(row=0, column=0, sticky="w")
        language_frame = ttk.Frame(frame)
        language_frame.grid(row=0, column=1, sticky="w")
        for column, tag in enumerate(["JA", "EN", "ZH"]):
            ttk.Checkbutton(
                language_frame, text=tag, variable=self.language_vars[tag]
            ).grid(row=0, column=column, padx=(0, 8), sticky="w")

        ttk.Label(frame, text="Model Size").grid(row=0, column=2, sticky="w")
        ttk.Combobox(
            frame,
            textvariable=self.model_size_var,
            values=["0.6B", "1.7B"],
            state="readonly",
            width=10,
        ).grid(row=0, column=3, sticky="w")

        ttk.Label(frame, text="Device").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.device_var).grid(
            row=1, column=1, sticky="ew", pady=(8, 0)
        )

        ttk.Label(frame, text="DType").grid(row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Combobox(
            frame,
            textvariable=self.dtype_var,
            values=["auto", "float16", "bfloat16", "float32"],
            state="readonly",
            width=10,
        ).grid(row=1, column=3, sticky="w", pady=(8, 0))

        ttk.Label(frame, text="Slide Width").grid(
            row=2, column=0, sticky="w", pady=(8, 0)
        )
        ttk.Entry(frame, textvariable=self.slide_width_var, width=12).grid(
            row=2, column=1, sticky="w", pady=(8, 0)
        )
        ttk.Label(frame, text="Slide Height").grid(
            row=2, column=2, sticky="w", pady=(8, 0)
        )
        ttk.Entry(frame, textvariable=self.slide_height_var, width=12).grid(
            row=2, column=3, sticky="w", pady=(8, 0)
        )

        ttk.Label(frame, text="Padding Sec").grid(
            row=3, column=0, sticky="w", pady=(8, 0)
        )
        ttk.Entry(frame, textvariable=self.slide_padding_var, width=12).grid(
            row=3, column=1, sticky="w", pady=(8, 0)
        )
        ttk.Label(frame, text="FPS").grid(row=3, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.fps_var, width=12).grid(
            row=3, column=3, sticky="w", pady=(8, 0)
        )

    def _build_actions_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, padding=(0, 12, 0, 0))
        frame.grid(row=2, column=0, sticky="ew")

        self.build_button = ttk.Button(
            frame, text="Build Videos", command=self.start_build_videos
        )
        self.build_button.grid(row=0, column=0, padx=(0, 8))

        self.output_button = ttk.Button(
            frame, text="Open Output Folder", command=self.open_output_folder
        )
        self.output_button.grid(row=0, column=1)

    def _build_status_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, padding=(0, 12, 0, 0))
        frame.grid(row=3, column=0, sticky="ew")
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=220)
        self.progress.grid(row=0, column=1, sticky="e")

    def _build_log_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Logs", padding=12)
        frame.grid(row=4, column=0, sticky="nsew", pady=(12, 0))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log_text = ScrolledText(frame, height=20, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

    def _add_file_row(
        self,
        parent: ttk.LabelFrame,
        *,
        row: int,
        label: str,
        variable: tk.StringVar,
        filetypes: list[tuple[str, str]],
        save_dialog: bool = False,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", pady=(8, 0), padx=(0, 8)
        )
        ttk.Button(
            parent,
            text="Browse",
            command=lambda: self._browse_file(variable, filetypes, save_dialog),
        ).grid(row=row, column=2, sticky="ew", pady=(8, 0))

    def _browse_file(
        self,
        variable: tk.StringVar,
        filetypes: list[tuple[str, str]],
        save_dialog: bool,
    ) -> None:
        initial_dir = str(self.project_root)
        if save_dialog:
            path = filedialog.asksaveasfilename(
                parent=self.root,
                initialdir=initial_dir,
                filetypes=filetypes,
                defaultextension=".pt",
            )
        else:
            path = filedialog.askopenfilename(
                parent=self.root,
                initialdir=initial_dir,
                filetypes=filetypes,
            )
        if path:
            variable.set(path)

    def _guess_pptx(self) -> str:
        for path in sorted(self.project_root.glob("*.pptx")):
            if not path.name.startswith("~$"):
                return str(path)
        return ""

    def _guess_voice_audio(self) -> str:
        for path in sorted(self.project_root.glob("*.wav")):
            if path.with_suffix(".txt").exists():
                return str(path)
        return ""

    def _guess_voice_text(self, audio: str | None = None) -> str:
        audio = audio or self._guess_voice_audio()
        if audio:
            return str(Path(audio).with_suffix(".txt"))
        return ""

    def _set_workspace_root(self, project_root: Path) -> None:
        self.project_root = project_root.resolve()
        self.paths = WorkspacePaths.from_root(self.project_root)
        self.paths.ensure_directories()

    def _selected_languages(self) -> list[str]:
        return [tag for tag, enabled in self.language_vars.items() if enabled.get()]

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.build_button.configure(state=state)
        if running:
            self.progress.start(12)
        else:
            self.progress.stop()

    def _poll_events(self) -> None:
        while True:
            try:
                event_type, payload = self.event_queue.get_nowait()
            except queue.Empty:
                break

            if event_type == "log":
                self._append_log(payload)
            elif event_type == "status":
                self.status_var.set(payload)
            elif event_type == "done":
                self.status_var.set(payload)
                self._set_running(False)
                messagebox.showinfo("Completed", payload, parent=self.root)
            elif event_type == "error":
                self.status_var.set("Failed")
                self._set_running(False)
                messagebox.showerror("Error", payload, parent=self.root)
        self.root.after(100, self._poll_events)

    def _require_build_inputs(self) -> dict[str, Any] | None:
        pptx_path = self._validated_path(self.pptx_var.get().strip(), "PPTX")
        if pptx_path is None:
            return None

        languages = self._selected_languages()
        if not languages:
            messagebox.showerror(
                "Missing Languages", "Select at least one language.", parent=self.root
            )
            return None

        voice_kwargs = self._voice_kwargs()
        if voice_kwargs is None:
            return None

        try:
            slide_width = int(self.slide_width_var.get().strip())
            slide_height = int(self.slide_height_var.get().strip())
            slide_padding_sec = float(self.slide_padding_var.get().strip())
            fps = int(self.fps_var.get().strip())
        except ValueError:
            messagebox.showerror(
                "Invalid Numeric Input",
                "Slide width, slide height, padding, and FPS must be numeric.",
                parent=self.root,
            )
            return None

        return {
            "pptx_path": pptx_path,
            "project_root": pptx_path.parent,
            "languages": languages,
            "model_size": self.model_size_var.get().strip(),
            "device": self._optional_text(self.device_var.get()),
            "dtype": self.dtype_var.get().strip(),
            "slide_width": slide_width,
            "slide_height": slide_height,
            "slide_padding_sec": slide_padding_sec,
            "fps": fps,
            **voice_kwargs,
        }

    def _voice_kwargs(self) -> dict[str, Any] | None:
        ref_audio_raw = self.ref_audio_var.get().strip()
        ref_text_raw = self.ref_text_var.get().strip()
        if not ref_audio_raw and not ref_text_raw:
            messagebox.showerror(
                "Missing Input",
                "Provide Ref Audio and Ref Text.",
                parent=self.root,
            )
            return None

        ref_audio = self._validated_path(ref_audio_raw, "Ref Audio")
        if ref_audio is None:
            return None
        ref_text = self._validated_path(ref_text_raw, "Ref Text")
        if ref_text is None:
            return None
        return {
            "voice_name": None,
            "ref_audio": ref_audio,
            "ref_text": ref_text,
        }

    def _validated_path(self, raw_value: str, label: str) -> Path | None:
        if not raw_value:
            messagebox.showerror(
                "Missing Input", f"{label} is required.", parent=self.root
            )
            return None
        path = Path(raw_value)
        if not path.exists():
            messagebox.showerror(
                "Missing File", f"{label} not found: {path}", parent=self.root
            )
            return None
        return path

    def _optional_text(self, value: str) -> str | None:
        stripped = value.strip()
        return stripped or None

    def start_build_videos(self) -> None:
        kwargs = self._require_build_inputs()
        if kwargs is None or self.worker_thread is not None:
            return
        self._set_workspace_root(cast(Path, kwargs["project_root"]))
        self.last_open_dir = self.paths.output
        logger.info("Starting video build...")
        self._run_worker(self._build_videos_worker, kwargs)

    def _run_worker(self, target, kwargs: dict[str, Any]) -> None:
        self._set_running(True)
        self.status_var.set("Working...")
        self.worker_thread = threading.Thread(
            target=target,
            kwargs=kwargs,
            daemon=True,
        )
        self.worker_thread.start()

    def _format_worker_error(self, exc: Exception) -> str:
        if isinstance(exc, ModuleNotFoundError) and exc.name == "_cffi_backend":
            return (
                "Missing Python dependency '_cffi_backend' required by soundfile. "
                "Install 'cffi' in the active environment or recreate the Conda environment."
            )
        if isinstance(exc, ModuleNotFoundError) and exc.name == "PIL.Image":
            return (
                "Pillow is installed incompletely or broken in the active environment. "
                "Reinstall 'pillow' or recreate the Conda environment."
            )
        return f"{type(exc).__name__}: {exc}"

    def _build_videos_worker(self, **kwargs: Any) -> None:
        try:
            logger.info(
                "GUI build request: pptx=%s, ref_audio=%s, ref_text=%s, languages=%s, model=%s, device=%s, dtype=%s",
                kwargs["pptx_path"],
                kwargs["ref_audio"],
                kwargs["ref_text"],
                ",".join(cast(list[str], kwargs["languages"])),
                kwargs["model_size"],
                kwargs["device"] or "auto",
                kwargs["dtype"],
            )
            from app.pipeline.build_movie import build_movie

            outputs = build_movie(
                pptx_path=cast(Path, kwargs["pptx_path"]),
                project_root=cast(Path, kwargs["project_root"]),
                voice_name=cast(str | None, kwargs["voice_name"]),
                ref_audio=cast(Path | None, kwargs["ref_audio"]),
                ref_text=cast(Path | None, kwargs["ref_text"]),
                languages=cast(list[str], kwargs["languages"]),
                model_size=cast(str, kwargs["model_size"]),
                device=cast(str | None, kwargs["device"]),
                dtype=cast(str | None, kwargs["dtype"]),
                slide_width=cast(int, kwargs["slide_width"]),
                slide_height=cast(int, kwargs["slide_height"]),
                slide_padding_sec=cast(float, kwargs["slide_padding_sec"]),
                fps=cast(int, kwargs["fps"]),
                prompt_cache_path=None,
                refresh_prompt_cache=False,
                use_prompt_cache=False,
                progress_callback=self._queue_status,
            )
            output_text = "\n".join(str(path) for path in outputs)
            self.event_queue.put(("done", f"Videos created:\n{output_text}"))
        except Exception as exc:
            logger.exception("Video build failed")
            self.event_queue.put(("error", self._format_worker_error(exc)))
        finally:
            self.worker_thread = None

    def _queue_status(self, message: str) -> None:
        self.event_queue.put(("status", message))

    def open_output_folder(self) -> None:
        output_dir = self.last_open_dir
        output_dir.mkdir(parents=True, exist_ok=True)
        try:
            if hasattr(os, "startfile"):
                os.startfile(output_dir)  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", str(output_dir)], check=True)
        except Exception as exc:
            messagebox.showerror(
                "Open Output Folder",
                f"Could not open {output_dir}: {exc}",
                parent=self.root,
            )


def main() -> None:
    root = tk.Tk()
    app = MovieBuilderGui(root)

    def on_close() -> None:
        root_logger = logging.getLogger()
        root_logger.removeHandler(app.log_handler)
        root_logger.removeHandler(app.stdout_handler)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()
