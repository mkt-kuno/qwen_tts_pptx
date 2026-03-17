from __future__ import annotations

import argparse
import logging
from collections.abc import Callable
from pathlib import Path

from app.common.paths import WorkspacePaths
from app.models.voices import resolve_voice_asset
from app.synthesis.qwen import detect_device, load_model, save_prompt_cache


logger = logging.getLogger(__name__)
ProgressCallback = Callable[[str], None]


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Export a reusable Qwen voice prompt cache."
    )
    parser.add_argument(
        "--voice-name",
        type=str,
        default=None,
        help="Voice base name, for example 'kuwano'",
    )
    parser.add_argument(
        "--ref-audio", type=Path, default=None, help="Reference WAV path"
    )
    parser.add_argument(
        "--ref-text", type=Path, default=None, help="Reference transcript path"
    )
    parser.add_argument("--model-size", choices=["0.6B", "1.7B"], default="1.7B")
    parser.add_argument("--device", type=str, default=None)
    parser.add_argument(
        "--dtype",
        type=str,
        default="auto",
        help="Torch dtype for model load: auto, bfloat16, float16, or float32",
    )
    parser.add_argument("--output", type=Path, default=None, help="Output .pt path")
    return parser


def export_voice_prompt(
    *,
    project_root: Path,
    voice_name: str | None,
    ref_audio: Path | None,
    ref_text: Path | None,
    model_size: str,
    device: str | None,
    dtype: str | None,
    output: Path | None,
    progress_callback: ProgressCallback | None = None,
) -> Path:
    voice = resolve_voice_asset(
        project_root=project_root,
        voice_name=voice_name,
        ref_audio=ref_audio,
        ref_text=ref_text,
    )
    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()

    resolved_device = detect_device(device)
    logger.info("Export workspace: %s", project_root)
    logger.info("Prompt output directory: %s", paths.prompts)
    logger.info("Using voice asset '%s'", voice.name)
    _report_progress(progress_callback, f"Loading model on {resolved_device}")
    model_result = load_model(
        model_size=model_size, device=resolved_device, dtype=dtype
    )
    model = model_result.model
    if model_result.device != resolved_device:
        resolved_device = model_result.device
        logger.warning("Model load fell back to %s", resolved_device)
        _report_progress(
            progress_callback, f"Model load fell back to {resolved_device}"
        )
    output_path = output or (paths.prompts / f"{voice.name}-{model_size}.pt")
    logger.info("Prompt cache target: %s", output_path)
    _report_progress(progress_callback, "Exporting prompt cache")
    save_prompt_cache(
        model=model,
        voice_name=voice.name,
        model_size=model_size,
        ref_audio=voice.audio_path,
        ref_text=voice.transcript,
        output_path=output_path,
    )
    logger.info("Voice prompt exported: %s", output_path)
    _report_progress(progress_callback, "Completed")
    return output_path


def _report_progress(progress_callback: ProgressCallback | None, message: str) -> None:
    if progress_callback is not None:
        progress_callback(message)


def main() -> None:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
    )
    args = build_arg_parser().parse_args()

    export_voice_prompt(
        project_root=Path.cwd(),
        voice_name=args.voice_name,
        ref_audio=args.ref_audio,
        ref_text=args.ref_text,
        model_size=args.model_size,
        device=args.device,
        dtype=args.dtype,
        output=args.output,
    )


if __name__ == "__main__":
    main()
