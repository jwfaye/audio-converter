"""CLI for converting between Excel and WAV formats.

Usage:
    audio-converter excel2wav input.xlsx output.wav --sample-rate 16000
    audio-converter wav2excel input.wav output.xlsx --sample-rate 16000
"""

import argparse
import sys
from pathlib import Path

from audio_converter.converter import ConversionError, excel_to_wav, wav_to_excel


def handle_excel2wav(args: argparse.Namespace) -> None:
    """CLI handler for Excel → WAV conversion."""
    try:
        result = excel_to_wav(args.input, args.output, args.sample_rate)
        print(f"WAV écrit : {result.output_path} ({result.num_samples} échantillons, {result.sample_rate} Hz)")
    except ConversionError as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        sys.exit(1)


def handle_wav2excel(args: argparse.Namespace) -> None:
    """CLI handler for WAV → Excel conversion."""
    try:
        result = wav_to_excel(args.input, args.output, args.sample_rate)
        print(f"Excel écrit : {result.output_path} ({result.num_samples} échantillons, {result.sample_rate} Hz)")
    except ConversionError as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        sys.exit(1)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convertir entre fichiers Excel et WAV",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # excel2wav
    p_e2w = subparsers.add_parser("excel2wav", help="Excel → WAV")
    p_e2w.add_argument("input", type=Path, help="Fichier Excel d'entrée (.xlsx)")
    p_e2w.add_argument("output", type=Path, help="Fichier WAV de sortie (.wav)")
    p_e2w.add_argument("--sample-rate", type=int, required=True, help="Fréquence d'échantillonnage (Hz)")
    p_e2w.set_defaults(func=handle_excel2wav)

    # wav2excel
    p_w2e = subparsers.add_parser("wav2excel", help="WAV → Excel")
    p_w2e.add_argument("input", type=Path, help="Fichier WAV d'entrée (.wav)")
    p_w2e.add_argument("output", type=Path, help="Fichier Excel de sortie (.xlsx)")
    p_w2e.add_argument("--sample-rate", type=int, default=None, help="Rééchantillonner à cette fréquence (Hz)")
    p_w2e.set_defaults(func=handle_wav2excel)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
