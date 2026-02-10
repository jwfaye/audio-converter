"""Tkinter GUI for audio-converter (Excel <-> WAV)."""

import re
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path

from tkinterdnd2 import DND_FILES, TkinterDnD

from audio_converter.converter import ConversionError, excel_to_wav, wav_to_excel

EXCEL_FILETYPES = [("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
WAV_FILETYPES = [("Fichiers WAV", "*.wav"), ("Tous les fichiers", "*.*")]


_SUPPORTED_EXTENSIONS = {".xlsx", ".wav"}


def _clean_drop_path(raw: str) -> str:
    """Normalize a path received from tkinterdnd2.

    On Windows, paths containing spaces are wrapped in braces: ``{C:/My Folder/file.wav}``.
    Multiple files may be space-separated; we only keep the first one.
    """
    raw = raw.strip()
    # Take only the first file if multiple were dropped
    m = re.match(r"\{([^}]+)\}", raw)
    if m:
        return m.group(1)
    return raw.split()[0] if raw else raw


class App(TkinterDnD.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Audio Converter")
        self.resizable(False, False)

        self._mode = tk.StringVar(value="excel2wav")
        self._input_path = tk.StringVar()
        self._output_path = tk.StringVar()
        self._sample_rate = tk.StringVar(value="16000")
        self._status = tk.StringVar(value="Prêt.")

        self._build_ui()
        self._on_mode_change()

    # ── UI construction ────────────────────────────────────────────

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 4}

        # Mode frame
        mode_frame = ttk.LabelFrame(self, text="Mode")
        mode_frame.grid(row=0, column=0, sticky="ew", **pad)

        ttk.Radiobutton(
            mode_frame, text="Excel → WAV", variable=self._mode,
            value="excel2wav", command=self._on_mode_change,
        ).grid(row=0, column=0, padx=12, pady=6)
        ttk.Radiobutton(
            mode_frame, text="WAV → Excel", variable=self._mode,
            value="wav2excel", command=self._on_mode_change,
        ).grid(row=0, column=1, padx=12, pady=6)

        # Drop zone
        self._drop_label = tk.Label(
            self,
            text="Glissez un fichier .xlsx ou .wav ici",
            relief="groove",
            bg="#e8f0fe",
            fg="#555555",
            font=("Segoe UI", 11),
            height=3,
            cursor="hand2",
        )
        self._drop_label.grid(row=1, column=0, sticky="ew", **pad)
        self._drop_label.drop_target_register(DND_FILES)
        self._drop_label.dnd_bind("<<Drop>>", self._on_drop)
        self._drop_label.dnd_bind("<<DragEnter>>", self._on_drag_enter)
        self._drop_label.dnd_bind("<<DragLeave>>", self._on_drag_leave)

        # Files frame
        files_frame = ttk.LabelFrame(self, text="Fichiers")
        files_frame.grid(row=2, column=0, sticky="ew", **pad)
        files_frame.columnconfigure(1, weight=1)

        ttk.Label(files_frame, text="Entrée :").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(files_frame, textvariable=self._input_path, width=50).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(files_frame, text="Parcourir", command=self._browse_input).grid(row=0, column=2, padx=4, pady=4)

        ttk.Label(files_frame, text="Sortie :").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(files_frame, textvariable=self._output_path, width=50).grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(files_frame, text="Parcourir", command=self._browse_output).grid(row=1, column=2, padx=4, pady=4)

        # Sample rate frame
        sr_frame = ttk.LabelFrame(self, text="Fréquence d'échantillonnage")
        sr_frame.grid(row=3, column=0, sticky="ew", **pad)

        ttk.Label(sr_frame, text="Hz :").grid(row=0, column=0, padx=4, pady=4)
        self._sr_entry = ttk.Entry(sr_frame, textvariable=self._sample_rate, width=10)
        self._sr_entry.grid(row=0, column=1, padx=4, pady=4)
        self._sr_hint = ttk.Label(sr_frame, text="")
        self._sr_hint.grid(row=0, column=2, padx=4, pady=4)

        # Convert button
        self._convert_btn = ttk.Button(self, text="Convertir", command=self._convert)
        self._convert_btn.grid(row=4, column=0, sticky="ew", **pad)

        # Status bar
        ttk.Label(self, textvariable=self._status, relief="sunken", anchor="w").grid(
            row=5, column=0, sticky="ew", padx=8, pady=(0, 8),
        )

    # ── Mode switching ─────────────────────────────────────────────

    def _on_mode_change(self) -> None:
        if self._mode.get() == "excel2wav":
            self._sr_hint.config(text="(requis)")
        else:
            self._sr_hint.config(text="(optionnel — rééchantillonnage)")

    # ── Drag & drop ────────────────────────────────────────────────

    def _on_drag_enter(self, event) -> None:
        self._drop_label.config(bg="#cfe2ff", fg="#003399")

    def _on_drag_leave(self, event) -> None:
        self._drop_label.config(bg="#e8f0fe", fg="#555555")

    def _on_drop(self, event) -> None:
        self._drop_label.config(bg="#e8f0fe", fg="#555555")
        path_str = _clean_drop_path(event.data)
        if not path_str:
            return

        p = Path(path_str)
        ext = p.suffix.lower()
        if ext not in _SUPPORTED_EXTENSIONS:
            self._status.set(f"Extension « {ext} » non supportée. Déposez un .xlsx ou .wav.")
            return

        # Auto-switch mode based on extension
        if ext == ".xlsx":
            self._mode.set("excel2wav")
        else:
            self._mode.set("wav2excel")
        self._on_mode_change()

        # Fill input and auto-generate output
        self._input_path.set(str(p))
        self._auto_output(str(p))
        self._status.set(f"Fichier chargé : {p.name}")

    # ── File browsing ──────────────────────────────────────────────

    def _browse_input(self) -> None:
        if self._mode.get() == "excel2wav":
            ft = EXCEL_FILETYPES
        else:
            ft = WAV_FILETYPES

        path = filedialog.askopenfilename(filetypes=ft)
        if path:
            self._input_path.set(path)
            self._auto_output(path)

    def _browse_output(self) -> None:
        if self._mode.get() == "excel2wav":
            ft = WAV_FILETYPES
        else:
            ft = EXCEL_FILETYPES

        path = filedialog.asksaveasfilename(filetypes=ft)
        if path:
            self._output_path.set(path)

    def _auto_output(self, input_path: str) -> None:
        """Generate an output path from the input path."""
        p = Path(input_path)
        if self._mode.get() == "excel2wav":
            self._output_path.set(str(p.with_suffix(".wav")))
        else:
            self._output_path.set(str(p.with_suffix(".xlsx")))

    # ── Validation ─────────────────────────────────────────────────

    def _validate(self) -> tuple[Path, Path, int | None] | None:
        """Validate inputs. Returns (input, output, sample_rate) or None."""
        input_str = self._input_path.get().strip()
        output_str = self._output_path.get().strip()
        sr_str = self._sample_rate.get().strip()

        if not input_str:
            self._status.set("Erreur : aucun fichier d'entrée sélectionné.")
            return None

        input_path = Path(input_str)
        if not input_path.exists():
            self._status.set(f"Erreur : fichier introuvable — {input_path}")
            return None

        if not output_str:
            self._status.set("Erreur : aucun fichier de sortie spécifié.")
            return None

        output_path = Path(output_str)

        mode = self._mode.get()
        if mode == "excel2wav":
            if not sr_str:
                self._status.set("Erreur : la fréquence d'échantillonnage est requise pour Excel → WAV.")
                return None
            try:
                sr = int(sr_str)
                if sr <= 0:
                    raise ValueError
            except ValueError:
                self._status.set("Erreur : fréquence d'échantillonnage invalide.")
                return None
        else:
            # WAV → Excel: sample rate is optional
            sr = None
            if sr_str:
                try:
                    sr = int(sr_str)
                    if sr <= 0:
                        raise ValueError
                except ValueError:
                    self._status.set("Erreur : fréquence d'échantillonnage invalide.")
                    return None

        return input_path, output_path, sr

    # ── Conversion (threaded) ──────────────────────────────────────

    def _convert(self) -> None:
        validated = self._validate()
        if validated is None:
            return

        input_path, output_path, sr = validated

        self._convert_btn.config(state="disabled")
        self._status.set("Conversion en cours...")

        thread = threading.Thread(
            target=self._run_conversion,
            args=(input_path, output_path, sr),
            daemon=True,
        )
        thread.start()

    def _run_conversion(self, input_path: Path, output_path: Path, sr: int | None) -> None:
        try:
            mode = self._mode.get()
            if mode == "excel2wav":
                result = excel_to_wav(input_path, output_path, sr)
            else:
                result = wav_to_excel(input_path, output_path, sr)

            msg = f"Terminé : {result.output_path.name} ({result.num_samples} échantillons, {result.sample_rate} Hz)"
            self.after(0, self._on_done, msg)

        except ConversionError as exc:
            self.after(0, self._on_done, f"Erreur : {exc}")
        except Exception as exc:
            self.after(0, self._on_done, f"Erreur inattendue : {exc}")

    def _on_done(self, message: str) -> None:
        self._status.set(message)
        self._convert_btn.config(state="normal")


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
