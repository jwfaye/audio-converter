"""Dark-themed Tkinter GUI for audio-converter with auditory pipeline."""

import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

from tkinterdnd2 import DND_FILES, TkinterDnD

from audio_converter.converter import (
    ConversionError,
    excel_to_wav,
    wav_to_pipeline_excel,
)

# ── Theme ─────────────────────────────────────────────────────────────────

C = {
    "bg": "#1a1a2e",
    "card": "#16213e",
    "border": "#0f3460",
    "text": "#e0e0e0",
    "dim": "#7a7a8e",
    "accent": "#2ecc71",
    "accent_h": "#27ae60",
    "white": "#ffffff",
    "drop": "#0f3460",
    "drop_h": "#1a4a7a",
    "drop_t": "#8899aa",
    "entry": "#0d1b2a",
    "sep": "#2a2a4e",
    "status": "#0d1b2a",
}

_SUPPORTED = {".xlsx", ".wav"}


def _clean_drop(raw: str) -> str:
    raw = raw.strip()
    m = re.match(r"\{([^}]+)\}", raw)
    return m.group(1) if m else (raw.split()[0] if raw else raw)


class App(TkinterDnD.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Audio Converter")
        self.configure(bg=C["bg"])
        self.resizable(False, False)

        # State
        self._input_path = tk.StringVar()
        self._output_path = tk.StringVar()
        self._sample_rate = tk.StringVar(value="16000")
        self._status = tk.StringVar(value="Prêt.")
        self._mode = "wav"

        # Pipeline options
        self._chk_audio = tk.BooleanVar(value=True)
        self._chk_periphery = tk.BooleanVar(value=True)
        self._chk_integration = tk.BooleanVar(value=False)
        self._tau = tk.StringVar(value="50")
        self._decimation = tk.StringVar(value="100")

        self._build()
        self._set_mode("wav")

    # ── Build ─────────────────────────────────────────────────────────

    def _build(self) -> None:
        # ── Drop zone ────────────────────────────────────────────────
        self._drop_frame = tk.Frame(self, bg=C["bg"])
        self._drop_frame.pack(fill="x", padx=16, pady=(16, 8))

        self._drop_zone = tk.Label(
            self._drop_frame,
            text="Glissez un fichier .wav ou .xlsx ici\n───\nou cliquez pour parcourir",
            bg=C["drop"],
            fg=C["drop_t"],
            font=("Helvetica", 11),
            cursor="hand2",
            pady=28,
            highlightbackground=C["border"],
            highlightthickness=1,
        )
        self._drop_zone.pack(fill="x")
        self._drop_zone.bind("<Button-1>", lambda e: self._browse_input())
        self._drop_zone.drop_target_register(DND_FILES)
        self._drop_zone.dnd_bind("<<Drop>>", self._on_drop)
        self._drop_zone.dnd_bind(
            "<<DragEnter>>", lambda e: self._drop_zone.config(bg=C["drop_h"])
        )
        self._drop_zone.dnd_bind(
            "<<DragLeave>>", lambda e: self._drop_zone.config(bg=C["drop"])
        )

        # ── Output section ───────────────────────────────────────────
        self._output_section, output_card = self._make_section("SORTIE")
        row = tk.Frame(output_card, bg=C["card"])
        row.pack(fill="x", padx=8, pady=8)

        tk.Entry(
            row,
            textvariable=self._output_path,
            bg=C["entry"],
            fg=C["text"],
            insertbackground=C["text"],
            relief="flat",
            font=("Helvetica", 10),
        ).pack(side="left", fill="x", expand=True, ipady=4)

        self._make_btn(row, "Parcourir", self._browse_output).pack(
            side="right", padx=(8, 0)
        )

        # ── Sample rate section (xlsx mode) ──────────────────────────
        self._sr_section, sr_card = self._make_section(
            "FRÉQUENCE D'ÉCHANTILLONNAGE"
        )
        sr_row = tk.Frame(sr_card, bg=C["card"])
        sr_row.pack(fill="x", padx=8, pady=8)

        tk.Label(
            sr_row,
            text="Hz :",
            bg=C["card"],
            fg=C["dim"],
            font=("Helvetica", 10),
        ).pack(side="left")

        tk.Entry(
            sr_row,
            textvariable=self._sample_rate,
            bg=C["entry"],
            fg=C["text"],
            insertbackground=C["text"],
            relief="flat",
            font=("Helvetica", 10),
            width=8,
        ).pack(side="left", padx=(8, 0), ipady=4)

        # ── Pipeline section (wav mode) ──────────────────────────────
        self._pipeline_section, pipeline_card = self._make_section(
            "PIPELINE PÉRIPHÉRIE AUDITIVE"
        )
        p = tk.Frame(pipeline_card, bg=C["card"])
        p.pack(fill="x", padx=8, pady=8)

        self._make_check(p, "Audio brut", self._chk_audio)
        self._make_sep(p)
        self._make_check(p, "Périphérie (memo MA + memo CK)", self._chk_periphery)
        self._make_sep(p)
        self._make_check(
            p,
            "Intégration temporelle",
            self._chk_integration,
            command=self._toggle_integration,
        )

        # Integration params (initially hidden)
        self._int_params = tk.Frame(p, bg=C["card"])
        self._make_param_row(self._int_params, "Tau", self._tau)
        self._make_param_row(self._int_params, "Décimation", self._decimation)

        # ── Convert button ───────────────────────────────────────────
        self._convert_btn = tk.Button(
            self,
            text="▶  CONVERTIR",
            bg=C["accent"],
            fg=C["white"],
            activebackground=C["accent_h"],
            activeforeground=C["white"],
            relief="flat",
            font=("Helvetica", 12, "bold"),
            cursor="hand2",
            command=self._convert,
            pady=10,
        )
        self._convert_btn.pack(fill="x", padx=16, pady=(8, 8))
        self._convert_btn.bind(
            "<Enter>", lambda e: self._convert_btn.config(bg=C["accent_h"])
        )
        self._convert_btn.bind(
            "<Leave>", lambda e: self._convert_btn.config(bg=C["accent"])
        )

        # ── Status bar ───────────────────────────────────────────────
        tk.Label(
            self,
            textvariable=self._status,
            bg=C["status"],
            fg=C["dim"],
            font=("Helvetica", 9),
            anchor="w",
            padx=8,
            pady=4,
        ).pack(fill="x", padx=16, pady=(0, 16))

    # ── Widget helpers ────────────────────────────────────────────────

    def _make_section(self, title: str) -> tuple[tk.Frame, tk.Frame]:
        container = tk.Frame(self, bg=C["bg"])
        container.pack(fill="x", padx=16, pady=4)

        tk.Label(
            container,
            text=title,
            bg=C["bg"],
            fg=C["dim"],
            font=("Helvetica", 9, "bold"),
            anchor="w",
        ).pack(fill="x", padx=4, pady=(0, 2))

        card = tk.Frame(
            container,
            bg=C["card"],
            highlightbackground=C["border"],
            highlightthickness=1,
        )
        card.pack(fill="x")

        return container, card

    def _make_btn(self, parent: tk.Widget, text: str, command) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            bg=C["border"],
            fg=C["text"],
            activebackground=C["drop_h"],
            activeforeground=C["text"],
            relief="flat",
            font=("Helvetica", 9),
            cursor="hand2",
            command=command,
        )
        btn.bind("<Enter>", lambda e: btn.config(bg=C["drop_h"]))
        btn.bind("<Leave>", lambda e: btn.config(bg=C["border"]))
        return btn

    def _make_check(
        self,
        parent: tk.Widget,
        text: str,
        variable: tk.BooleanVar,
        command=None,
    ) -> tk.Checkbutton:
        cb = tk.Checkbutton(
            parent,
            text=text,
            variable=variable,
            bg=C["card"],
            fg=C["text"],
            selectcolor=C["entry"],
            activebackground=C["card"],
            activeforeground=C["text"],
            font=("Helvetica", 10),
            anchor="w",
            command=command,
        )
        cb.pack(fill="x", pady=2)
        return cb

    def _make_sep(self, parent: tk.Widget) -> None:
        tk.Frame(parent, bg=C["sep"], height=1).pack(fill="x", pady=4)

    def _make_param_row(
        self, parent: tk.Widget, label: str, variable: tk.StringVar
    ) -> None:
        row = tk.Frame(parent, bg=C["card"])
        row.pack(fill="x", pady=2)

        tk.Label(
            row,
            text=label,
            bg=C["card"],
            fg=C["dim"],
            font=("Helvetica", 9),
            width=14,
            anchor="w",
        ).pack(side="left", padx=(24, 0))

        tk.Entry(
            row,
            textvariable=variable,
            bg=C["entry"],
            fg=C["text"],
            insertbackground=C["text"],
            relief="flat",
            font=("Helvetica", 10),
            width=8,
        ).pack(side="left", ipady=3)

    # ── Mode switching ────────────────────────────────────────────────

    def _set_mode(self, mode: str) -> None:
        self._mode = mode
        if mode == "wav":
            self._sr_section.pack_forget()
            self._pipeline_section.pack(
                after=self._output_section, fill="x", padx=16, pady=4
            )
        else:
            self._pipeline_section.pack_forget()
            self._sr_section.pack(
                after=self._output_section, fill="x", padx=16, pady=4
            )
        self._resize()

    def _toggle_integration(self) -> None:
        if self._chk_integration.get():
            self._int_params.pack(fill="x", pady=(4, 0))
        else:
            self._int_params.pack_forget()
        self._resize()

    def _resize(self) -> None:
        self.update_idletasks()
        self.geometry("")

    # ── Drag & drop ───────────────────────────────────────────────────

    def _on_drop(self, event) -> None:
        self._drop_zone.config(bg=C["drop"])
        path_str = _clean_drop(event.data)
        if not path_str:
            return

        p = Path(path_str)
        ext = p.suffix.lower()
        if ext not in _SUPPORTED:
            self._status.set(
                f"Extension \u00ab {ext} \u00bb non support\u00e9e. "
                "D\u00e9posez un .xlsx ou .wav."
            )
            return

        self._input_path.set(str(p))
        self._auto_output(p)

        if ext == ".xlsx":
            self._set_mode("xlsx")
        else:
            self._set_mode("wav")

        self._status.set(f"Fichier charg\u00e9 : {p.name}")

    # ── File browsing ─────────────────────────────────────────────────

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[
                ("Audio / Excel", "*.wav *.xlsx"),
                ("Tous les fichiers", "*.*"),
            ]
        )
        if not path:
            return

        p = Path(path)
        self._input_path.set(str(p))
        self._auto_output(p)

        if p.suffix.lower() == ".xlsx":
            self._set_mode("xlsx")
        else:
            self._set_mode("wav")

        self._status.set(f"Fichier charg\u00e9 : {p.name}")

    def _browse_output(self) -> None:
        if self._mode == "xlsx":
            ft = [("Fichiers WAV", "*.wav"), ("Tous", "*.*")]
        else:
            ft = [("Fichiers Excel", "*.xlsx"), ("Tous", "*.*")]

        path = filedialog.asksaveasfilename(filetypes=ft)
        if path:
            self._output_path.set(path)

    def _auto_output(self, p: Path) -> None:
        if p.suffix.lower() == ".xlsx":
            self._output_path.set(str(p.with_suffix(".wav")))
        else:
            self._output_path.set(str(p.with_suffix(".xlsx")))

    # ── Validation ────────────────────────────────────────────────────

    def _validate(self) -> tuple[Path, Path] | None:
        input_str = self._input_path.get().strip()
        output_str = self._output_path.get().strip()

        if not input_str:
            self._status.set("Erreur : aucun fichier d'entr\u00e9e s\u00e9lectionn\u00e9.")
            return None

        input_path = Path(input_str)
        if not input_path.exists():
            self._status.set(f"Erreur : fichier introuvable \u2014 {input_path}")
            return None

        if not output_str:
            self._status.set("Erreur : aucun fichier de sortie sp\u00e9cifi\u00e9.")
            return None

        return input_path, Path(output_str)

    # ── Conversion (threaded) ─────────────────────────────────────────

    def _convert(self) -> None:
        validated = self._validate()
        if validated is None:
            return

        input_path, output_path = validated

        if self._mode == "xlsx":
            sr_str = self._sample_rate.get().strip()
            if not sr_str:
                self._status.set(
                    "Erreur : la fr\u00e9quence d'\u00e9chantillonnage est requise."
                )
                return
            try:
                sr = int(sr_str)
                if sr <= 0:
                    raise ValueError
            except ValueError:
                self._status.set(
                    "Erreur : fr\u00e9quence d'\u00e9chantillonnage invalide."
                )
                return

        self._convert_btn.config(state="disabled")
        self._status.set("Conversion en cours...")

        thread = threading.Thread(
            target=self._run_conversion,
            args=(input_path, output_path),
            daemon=True,
        )
        thread.start()

    def _run_conversion(self, input_path: Path, output_path: Path) -> None:
        def on_progress(msg: str) -> None:
            self.after(0, self._status.set, msg)

        try:
            if self._mode == "xlsx":
                sr = int(self._sample_rate.get())
                result = excel_to_wav(input_path, output_path, sr)
            else:
                tau = int(self._tau.get()) if self._tau.get().strip() else 50
                dec = (
                    int(self._decimation.get())
                    if self._decimation.get().strip()
                    else 100
                )
                result = wav_to_pipeline_excel(
                    input_path,
                    output_path,
                    export_audio=self._chk_audio.get(),
                    export_periphery=self._chk_periphery.get(),
                    export_integration=self._chk_integration.get(),
                    tau=tau,
                    decimation_factor=dec,
                    progress_callback=on_progress,
                )

            msg = (
                f"Termin\u00e9 : {result.output_path.name} "
                f"({result.num_samples} \u00e9chantillons, {result.sample_rate} Hz)"
            )
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
