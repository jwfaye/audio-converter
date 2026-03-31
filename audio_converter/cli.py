"""Command-line interface for converting between Excel and WAV formats.

Sub-commands::

    audio-converter excel2wav input.xlsx output.wav --sample-rate 16000
    audio-converter wav2excel input.wav output.xlsx [--sample-rate 16000]
    audio-converter pipeline  input.wav output.xlsx [OPTIONS]

The ``pipeline`` sub-command runs the auditory periphery pipeline and
exports selectable stages (audio, peaks, memo_ma/ck, integration) to a
multi-sheet Excel workbook.
"""

import argparse
import sys
from pathlib import Path

from audio_converter.converter import (
    ConversionError,
    excel_to_wav,
    wav_to_excel,
    wav_to_pipeline_excel,
)


def handle_excel2wav(args: argparse.Namespace) -> None:
    """Handle the ``excel2wav`` sub-command.

    Reads int16 samples from the first row of each ``audio*`` sheet and
    writes a PCM-16 WAV file.

    Args:
        args: Parsed CLI arguments (``input``, ``output``,
            ``sample_rate``).
    """
    try:
        result = excel_to_wav(args.input, args.output, args.sample_rate)
        print(
            f"WAV écrit : {result.output_path} ({result.num_samples} échantillons, {result.sample_rate} Hz)"
        )
    except ConversionError as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        sys.exit(1)


def handle_wav2excel(args: argparse.Namespace) -> None:
    """Handle the ``wav2excel`` sub-command.

    Loads a WAV file, resamples to 16 kHz, and writes int16 samples
    to a single-row Excel sheet.

    Args:
        args: Parsed CLI arguments (``input``, ``output``).
    """
    try:
        result = wav_to_excel(args.input, args.output)
        print(
            f"Excel écrit : {result.output_path} ({result.num_samples} échantillons, {result.sample_rate} Hz)"
        )
    except ConversionError as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        sys.exit(1)


def handle_pipeline(args: argparse.Namespace) -> None:
    """Handle the ``pipeline`` sub-command.

    Runs the auditory periphery pipeline on a WAV file and writes the
    selected stages to a multi-sheet Excel workbook.

    Args:
        args: Parsed CLI arguments (``input``, ``output``,
            ``no_audio``, ``no_periphery``, ``integration``, ``tau``,
            ``decimation``).
    """
    try:
        result = wav_to_pipeline_excel(
            args.input,
            args.output,
            export_audio=not args.no_audio,
            export_periphery=not args.no_periphery,
            export_integration=args.integration,
            tau=args.tau,
            decimation_factor=args.decimation,
            progress_callback=lambda msg: print(f"  {msg}"),
        )
        print(f"Excel écrit : {result.output_path} ({result.num_samples} échantillons, {result.sample_rate} Hz)")
    except ConversionError as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        sys.exit(1)


def main() -> None:
    """Entry point for the ``audio-converter`` CLI.

    Parses arguments and dispatches to the appropriate sub-command
    handler.
    """
    parser = argparse.ArgumentParser(
        description="Convertir entre fichiers Excel et WAV",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    p_e2w = subparsers.add_parser("excel2wav", help="Excel → WAV")
    p_e2w.add_argument("input", type=Path, help="Fichier Excel d'entrée (.xlsx)")
    p_e2w.add_argument("output", type=Path, help="Fichier WAV de sortie (.wav)")
    p_e2w.add_argument(
        "--sample-rate",
        type=int,
        required=True,
        help="Fréquence d'échantillonnage (Hz)",
    )
    p_e2w.set_defaults(func=handle_excel2wav)

    p_w2e = subparsers.add_parser("wav2excel", help="WAV → Excel (rééchantillonné à 16 kHz)")
    p_w2e.add_argument("input", type=Path, help="Fichier WAV d'entrée (.wav)")
    p_w2e.add_argument("output", type=Path, help="Fichier Excel de sortie (.xlsx)")
    p_w2e.set_defaults(func=handle_wav2excel)

    p_pipe = subparsers.add_parser("pipeline", help="WAV → Excel multi-feuilles (pipeline périphérie)")
    p_pipe.add_argument("input", type=Path, help="Fichier WAV d'entrée (.wav)")
    p_pipe.add_argument("output", type=Path, help="Fichier Excel de sortie (.xlsx)")
    p_pipe.add_argument("--no-audio", action="store_true", help="Ne pas exporter l'audio brut")
    p_pipe.add_argument("--no-periphery", action="store_true", help="Ne pas exporter memo MA/CK")
    p_pipe.add_argument("--integration", action="store_true", help="Exporter l'intégration temporelle")
    p_pipe.add_argument("--tau", type=int, default=200, help="Constante d'intégration temporelle (défaut: 200)")
    p_pipe.add_argument("--decimation", type=int, default=100, help="Facteur de décimation (défaut: 100)")
    p_pipe.set_defaults(func=handle_pipeline)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
