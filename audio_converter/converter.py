"""Business logic for converting between Excel and WAV formats.

Shared by CLI and GUI — no dependency on argparse, print, or sys.exit().
"""

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

import numpy as np
from openpyxl import Workbook, load_workbook

from audio_converter.audio_io import load_audio, save_audio

INT16_MAX = 32768.0
EXCEL_MAX_COLS = 16384


class ConversionError(Exception):
    """Raised when a conversion fails."""


@dataclass
class ConversionResult:
    """Result of a successful conversion."""

    output_path: Path
    num_samples: int
    sample_rate: int


def excel_to_wav(
    input_path: Path,
    output_path: Path,
    sample_rate: int,
) -> ConversionResult:
    """Convert an Excel file (first row of samples) to a WAV file.

    Args:
        input_path: Path to the input .xlsx file.
        output_path: Path to the output .wav file.
        sample_rate: Sample rate in Hz.

    Returns:
        ConversionResult with details of the written file.

    Raises:
        ConversionError: If no data is found or conversion fails.
    """
    try:
        wb = load_workbook(input_path, read_only=True, data_only=True)
    except Exception as exc:
        raise ConversionError(f"Impossible d'ouvrir le fichier Excel : {exc}") from exc

    audio_sheets = sorted(
        [s for s in wb.sheetnames if s.startswith("audio")],
        key=lambda s: (int(s.split("_")[1]) if "_" in s else 0),
    )
    if not audio_sheets:
        audio_sheets = [wb.sheetnames[0]]

    samples: list[float] = []
    for sheet_name in audio_sheets:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            for val in row:
                if val is not None:
                    samples.append(float(val))
    wb.close()

    if not samples:
        raise ConversionError(
            "Aucune donnée trouvée dans la première ligne du fichier Excel."
        )

    audio_data = np.array(samples, dtype=np.float32)
    audio_data = audio_data / INT16_MAX

    save_audio(output_path, audio_data, sample_rate)

    return ConversionResult(
        output_path=output_path,
        num_samples=len(samples),
        sample_rate=sample_rate,
    )


def wav_to_excel(
    input_path: Path,
    output_path: Path,
    sample_rate: int | None = None,
) -> ConversionResult:
    """Convert a WAV file to an Excel file (one row of samples).

    Args:
        input_path: Path to the input .wav file.
        output_path: Path to the output .xlsx file.
        sample_rate: Optional target sample rate for resampling (Hz).

    Returns:
        ConversionResult with details of the written file.

    Raises:
        ConversionError: If the WAV file cannot be read or conversion fails.
    """
    try:
        audio_data, sr = load_audio(input_path, target_sample_rate=sample_rate)
    except Exception as exc:
        raise ConversionError(f"Impossible de lire le fichier WAV : {exc}") from exc

    wb = Workbook()
    wb.remove(wb.active)

    int16_data = (audio_data * INT16_MAX).astype(np.int16)
    _write_audio_sheets(wb, int16_data)

    try:
        wb.save(output_path)
    except Exception as exc:
        raise ConversionError(f"Impossible d'écrire le fichier Excel : {exc}") from exc

    return ConversionResult(
        output_path=output_path,
        num_samples=len(audio_data),
        sample_rate=sr,
    )


def wav_to_pipeline_excel(
    input_path: Path,
    output_path: Path,
    export_audio: bool = True,
    export_periphery: bool = True,
    export_integration: bool = False,
    tau: int = 50,
    decimation_factor: int = 100,
    progress_callback: Callable[[str], None] | None = None,
) -> ConversionResult:
    """Convert a WAV file to a multi-sheet Excel with auditory pipeline stages.

    Each enabled stage produces one or more sheets:
    - audio: raw signal (1 row)
    - memo_ma / memo_ck: auditory nerve output (65 channels x n_samples)
    - memo_ma_int / memo_ck_int: after temporal integration (65 x n_decimated)

    Args:
        input_path: Path to the input .wav file.
        output_path: Path to the output .xlsx file.
        export_audio: Include raw audio sheet.
        export_periphery: Include memo_ma and memo_ck sheets.
        export_integration: Include integrated memo_ma and memo_ck sheets.
        tau: Temporal integration time constant (default: 50).
        decimation_factor: Decimation factor (default: 100).
        progress_callback: Optional callback for progress messages.

    Returns:
        ConversionResult with details of the written file.

    Raises:
        ConversionError: If conversion fails or soundperception is not installed.
    """
    def _report(msg: str) -> None:
        if progress_callback:
            progress_callback(msg)

    try:
        from soundperception.audition.config import (
            AuditoryPeripheryConfig,
            TemporalIntegrationConfig,
        )
        from soundperception.audition.core.integration import TemporalIntegrator
        from soundperception.audition.core.periphery import AuditoryPeriphery
    except ImportError as exc:
        raise ConversionError(
            "Le module soundperception n'est pas installé. "
            "Installez-le avec : pip install -e ../sound-perception"
        ) from exc

    _report("Chargement audio...")
    try:
        audio_data, sr = load_audio(input_path)
    except Exception as exc:
        raise ConversionError(f"Impossible de lire le fichier WAV : {exc}") from exc

    wb = Workbook()
    wb.remove(wb.active)

    if export_audio:
        _report("Export audio brut...")
        int16_data = (audio_data * INT16_MAX).astype(np.int16)
        _write_audio_sheets(wb, int16_data)

    needs_periphery = export_periphery or export_integration
    if needs_periphery:
        _report("Calcul périphérie auditive...")
        config = AuditoryPeripheryConfig.default()
        config.cochlear.sample_rate = sr
        periphery = AuditoryPeriphery(config)
        result = periphery.process(audio_data)

        memo_ma = result["memo_ma"]
        memo_ck = result["memo_ck"]

        if export_periphery:
            _report("Export memo MA...")
            ws_ma = wb.create_sheet("memo_ma")
            _write_2d_array(ws_ma, memo_ma)

            _report("Export memo CK...")
            ws_ck = wb.create_sheet("memo_ck")
            _write_2d_array(ws_ck, memo_ck)

        if export_integration:
            _report("Calcul intégration temporelle...")
            int_config = TemporalIntegrationConfig(
                tau=tau,
                decimation_factor=decimation_factor,
            )
            integrator = TemporalIntegrator(int_config)
            int_result = integrator.process(memo_ma, memo_ck)

            _report("Export memo MA intégré...")
            ws_ma_int = wb.create_sheet("memo_ma_int")
            _write_2d_array(ws_ma_int, int_result["memo_ma"])

            _report("Export memo CK intégré...")
            ws_ck_int = wb.create_sheet("memo_ck_int")
            _write_2d_array(ws_ck_int, int_result["memo_ck"])

    if not wb.sheetnames:
        raise ConversionError("Aucune étape sélectionnée pour l'export.")

    _report("Écriture fichier Excel...")
    try:
        wb.save(output_path)
    except Exception as exc:
        raise ConversionError(f"Impossible d'écrire le fichier Excel : {exc}") from exc

    return ConversionResult(
        output_path=output_path,
        num_samples=len(audio_data),
        sample_rate=sr,
    )


def _write_audio_sheets(wb: Workbook, int16_data: np.ndarray) -> None:
    """Write int16 audio samples across one or more 'audio' sheets.

    Each sheet holds up to EXCEL_MAX_COLS (16 384) samples in row 1.
    Sheets are named audio, audio_2, audio_3, etc.
    """
    chunks = [
        int16_data[i : i + EXCEL_MAX_COLS]
        for i in range(0, len(int16_data), EXCEL_MAX_COLS)
    ]
    for idx, chunk in enumerate(chunks):
        name = "audio" if idx == 0 else f"audio_{idx + 1}"
        ws = wb.create_sheet(name)
        ws.append([int(s) for s in chunk])


def _write_2d_array(ws, array: np.ndarray) -> None:
    """Write a 2D numpy array to a worksheet. Rows = channels, cols = samples."""
    for row in array:
        ws.append([float(v) for v in row])
