"""Audio I/O operations using soundfile with libsndfile backend.

Adapted from soundperception/utils/audio_io.py
"""

from pathlib import Path

import numpy as np
import soundfile as sf
from scipy import signal as scipy_signal


def load_audio(
    file_path: str | Path,
    target_sample_rate: int | None = None,
) -> tuple[np.ndarray, int]:
    """Load a WAV file and return mono float32 samples in [-1, 1].

    Args:
        file_path: Path to audio file.
        target_sample_rate: Target sample rate (Hz). If None, use original.

    Returns:
        (audio_data, sample_rate) where audio_data is float32 in [-1, 1].
    """
    audio_data, original_sr = sf.read(str(file_path), always_2d=False)

    if audio_data.ndim > 1:
        audio_data = np.mean(audio_data, axis=1)

    audio_data = audio_data.astype(np.float32)

    if target_sample_rate is not None and original_sr != target_sample_rate:
        num_samples = int(len(audio_data) * target_sample_rate / original_sr)
        audio_data = scipy_signal.resample(audio_data, num_samples).astype(np.float32)
        return audio_data, target_sample_rate

    return audio_data, original_sr


def save_audio(
    file_path: str | Path,
    audio_data: np.ndarray,
    sample_rate: int,
) -> None:
    """Save audio samples to a WAV file.

    Args:
        file_path: Output file path.
        audio_data: float32 or float64 samples in [-1, 1].
        sample_rate: Sample rate in Hz.
    """
    audio_data = np.asarray(audio_data, dtype=np.float32)

    max_val = np.abs(audio_data).max()
    if max_val > 1.0:
        audio_data = audio_data / max_val

    sf.write(str(file_path), audio_data, sample_rate, subtype="PCM_16")
