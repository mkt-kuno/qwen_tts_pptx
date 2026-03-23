from __future__ import annotations

import argparse
import logging
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from app.common.languages import LanguageSpec, resolve_languages
from app.common.notes import extract_tagged_text, iter_slide_notes
from app.common.paths import WorkspacePaths
from app.models.voices import VoiceAsset, resolve_voice_asset
from app.synthesis.qwen import (
    create_voice_clone_prompt,
    detect_device,
    load_model,
    load_prompt_cache,
    save_prompt_cache,
    synthesize_to_file,
)


logger = logging.getLogger(__name__)
ProgressCallback = Callable[[str], None]


@dataclass(frozen=True)
class AudioGenerationResult:
    slide_count: int
    active_languages: list[LanguageSpec]


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate language WAVs from PPTX notes with Qwen TTS."
    )
    parser.add_argument("pptx", type=Path, help="Input .pptx path")
    parser.add_argument("--voice-name", type=str, default=None)
    parser.add_argument("--ref-audio", type=Path, default=None)
    parser.add_argument("--ref-text", type=Path, default=None)
    parser.add_argument(
        "--languages",
        nargs="+",
        default=["EN", "JP", "ZH", "ES", "IT", "FR"],
    )
    parser.add_argument("--model-size", choices=["0.6B", "1.7B"], default="1.7B")
    parser.add_argument("--device", type=str, default=None)
    parser.add_argument(
        "--dtype",
        type=str,
        default="auto",
        help="Torch dtype for model load: auto, bfloat16, float16, or float32",
    )
    parser.add_argument("--prompt-cache", type=Path, default=None)
    parser.add_argument("--refresh-prompt-cache", action="store_true")
    return parser


def generate_audio_tracks(
    pptx_path: Path,
    voice: VoiceAsset,
    languages: list[LanguageSpec],
    paths: WorkspacePaths,
    model_size: str,
    device: str | None,
    dtype: str | None,
    prompt_cache_path: Path | None,
    refresh_prompt_cache: bool,
    use_prompt_cache: bool = True,
    progress_callback: ProgressCallback | None = None,
) -> AudioGenerationResult:
    notes = iter_slide_notes(pptx_path)
    slide_texts: list[tuple[int, dict[str, str]]] = []
    active_tags: set[str] = set()

    logger.info("Parsing slide notes from %s", pptx_path)
    logger.info(
        "Requested audio languages: %s",
        ", ".join(spec.tag for spec in languages),
    )

    for note in notes:
        tagged = {
            spec.tag: extract_tagged_text(note.raw_text, spec.tag) for spec in languages
        }
        for tag, text in tagged.items():
            if text.strip():
                active_tags.add(tag)
        slide_texts.append((note.slide_number, tagged))

    active_languages = [spec for spec in languages if spec.tag in active_tags]
    skipped = [spec.tag for spec in languages if spec.tag not in active_tags]
    logger.info("Found notes for %d slides", len(notes))
    for tag in skipped:
        logger.warning("No [%s] notes found. Skipping that output.", tag)

    if not active_languages:
        raise RuntimeError("No matching note tags were found in the presentation.")

    logger.info(
        "Active audio languages: %s",
        ", ".join(spec.tag for spec in active_languages),
    )

    for spec in active_languages:
        output_dir = paths.audio_dir(spec.directory_name)
        output_dir.mkdir(parents=True, exist_ok=True)
        logger.info("Preparing audio directory %s", output_dir)
        for wav_path in output_dir.glob("page*.wav"):
            wav_path.unlink()

    logger.info("Detecting inference device")
    resolved_device = detect_device(device)
    logger.info("Using inference device %s", resolved_device)
    _report_progress(progress_callback, f"Loading model on {resolved_device}")
    logger.info("Loading Qwen TTS model on %s", resolved_device)
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
    logger.info("Qwen TTS model loaded on %s", resolved_device)

    if not use_prompt_cache:
        _report_progress(progress_callback, "Building voice prompt")
        logger.info("Building voice prompt directly from %s", voice.audio_path)
        prompt = create_voice_clone_prompt(
            model=model,
            ref_audio=voice.audio_path,
            ref_text=voice.transcript,
        )
        logger.info("Voice prompt created")
    else:
        cache_path = prompt_cache_path or (
            paths.prompts / f"{voice.name}-{model_size}.pt"
        )
        if cache_path.exists() and not refresh_prompt_cache:
            _report_progress(progress_callback, "Loading prompt cache")
            logger.info("Loading prompt cache %s", cache_path)
            prompt = load_prompt_cache(cache_path)
            logger.info("Prompt cache loaded")
        else:
            _report_progress(progress_callback, "Building prompt cache")
            logger.info("Building prompt cache %s", cache_path)
            prompt = save_prompt_cache(
                model=model,
                voice_name=voice.name,
                model_size=model_size,
                ref_audio=voice.audio_path,
                ref_text=voice.transcript,
                output_path=cache_path,
            )
            logger.info("Prompt cache created at %s", cache_path)

    for slide_number, tagged in slide_texts:
        for spec in active_languages:
            _report_progress(
                progress_callback,
                f"Generating {spec.tag} audio for slide {slide_number}/{len(notes)}",
            )
            output_path = (
                paths.audio_dir(spec.directory_name) / f"page{slide_number}.wav"
            )
            logger.info(
                "Generating %s audio for slide %s -> %s",
                spec.tag,
                slide_number,
                output_path,
            )
            synthesize_to_file(
                model=model,
                prompt=prompt,
                text=tagged[spec.tag],
                language=spec.qwen_language,
                output_path=output_path,
            )

    return AudioGenerationResult(
        slide_count=len(notes), active_languages=active_languages
    )


def _report_progress(progress_callback: ProgressCallback | None, message: str) -> None:
    if progress_callback is not None:
        progress_callback(message)


def main() -> None:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
    )
    args = build_arg_parser().parse_args()

    project_root = Path.cwd()
    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()
    voice = resolve_voice_asset(
        project_root=project_root,
        voice_name=args.voice_name,
        ref_audio=args.ref_audio,
        ref_text=args.ref_text,
    )

    result = generate_audio_tracks(
        pptx_path=args.pptx,
        voice=voice,
        languages=resolve_languages(args.languages),
        paths=paths,
        model_size=args.model_size,
        device=args.device,
        dtype=args.dtype,
        prompt_cache_path=args.prompt_cache,
        refresh_prompt_cache=args.refresh_prompt_cache,
    )
    logger.info("Generated audio for %d slides", result.slide_count)


if __name__ == "__main__":
    main()
