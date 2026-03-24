"""Audio I/O helpers backed by *soundfile* (libsndfile).

Provides :func:`load_audio` and :func:`save_audio` for reading/writing
WAV files as float32 arrays normalised to ``[-1, 1]``.

Adapted from ``soundperception/utils/audio_io.py``.
"""

from pathlib import Path

import numpy as np
import soundfile as sf
from scipy import signal as scipy_signal


def load_audio(
    file_path: str | Path,
    target_sample_rate: int | None = None,
) -> tuple[np.ndarray, int]:
    """Load a WAV file and return mono float32 samples in ``[-1, 1]``.

    Multi-channel files are down-mixed to mono by averaging.  If
    *target_sample_rate* differs from the file's native rate, the signal
    is resampled with :func:`scipy.signal.resample`.

    Args:
        file_path: Path to the audio file (WAV, FLAC, OGG, ...).
        target_sample_rate: Desired sample rate in Hz.  ``None`` keeps
            the original rate.

    Returns:
        A tuple ``(audio_data, sample_rate)`` where *audio_data* is a
        1-D float32 array in ``[-1, 1]``.
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
    """Save audio samples to a 16-bit PCM WAV file.

    Args:
        file_path: Destination path (parent directory must exist).
        audio_data: Float32 or float64 samples in ``[-1, 1]``.
        sample_rate: Sample rate in Hz.
    """
    audio_data = np.asarray(audio_data, dtype=np.float32)
    sf.write(str(file_path), audio_data, sample_rate, subtype="PCM_16")
