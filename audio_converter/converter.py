"""Business logic for converting between Excel and WAV formats.

Shared by CLI and GUI — no dependency on argparse, print, or sys.exit().
"""

from dataclasses import dataclass
from pathlib import Path

import numpy as np
from openpyxl import Workbook, load_workbook

from audio_converter.audio_io import load_audio, save_audio


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

    ws = wb.active
    samples: list[float] = []
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

    if np.abs(audio_data).max() > 1.0:
        audio_data = audio_data / np.abs(audio_data).max()

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
    ws = wb.active
    ws.title = "audio"

    for col_idx, sample in enumerate(audio_data, start=1):
        ws.cell(row=1, column=col_idx, value=float(sample))

    try:
        wb.save(output_path)
    except Exception as exc:
        raise ConversionError(f"Impossible d'écrire le fichier Excel : {exc}") from exc

    return ConversionResult(
        output_path=output_path,
        num_samples=len(audio_data),
        sample_rate=sr,
    )
