from __future__ import annotations

import argparse
import logging
from collections.abc import Sequence
from collections.abc import Callable
from pathlib import Path

from app.common.languages import LanguageSpec, resolve_languages
from app.common.paths import WorkspacePaths
from app.models.voices import resolve_voice_asset
from app.slides.export_slides import export_slides
from app.synthesis.generate_audio import generate_audio_tracks
from app.video.build import build_videos


logger = logging.getLogger(__name__)
ProgressCallback = Callable[[str], None]


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build language-specific MP4s from a PowerPoint deck and Qwen TTS voice assets."
    )
    parser.add_argument("pptx", type=Path, help="Input .pptx path")
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
    parser.add_argument(
        "--languages",
        nargs="+",
        default=["JA", "EN", "ZH"],
        help="Note tags to process",
    )
    parser.add_argument("--model-size", choices=["0.6B", "1.7B"], default="0.6B")
    parser.add_argument(
        "--device", type=str, default=None, help="Torch device, for example 'cuda:0'"
    )
    parser.add_argument(
        "--dtype",
        type=str,
        default="auto",
        help="Torch dtype for model load: auto, bfloat16, float16, or float32",
    )
    parser.add_argument("--slide-width", type=int, default=2560)
    parser.add_argument("--slide-height", type=int, default=1440)
    parser.add_argument("--slide-padding-sec", type=float, default=1.5)
    parser.add_argument("--fps", type=int, default=5)
    parser.add_argument(
        "--prompt-cache", type=Path, default=None, help="Optional .pt prompt cache path"
    )
    parser.add_argument(
        "--refresh-prompt-cache", action="store_true", help="Rebuild the prompt cache"
    )
    return parser


def build_movie(
    *,
    pptx_path: Path,
    project_root: Path,
    voice_name: str | None,
    ref_audio: Path | None,
    ref_text: Path | None,
    languages: Sequence[str | LanguageSpec],
    model_size: str,
    device: str | None,
    dtype: str | None,
    slide_width: int,
    slide_height: int,
    slide_padding_sec: float,
    fps: int,
    prompt_cache_path: Path | None,
    refresh_prompt_cache: bool,
    use_prompt_cache: bool = True,
    progress_callback: ProgressCallback | None = None,
) -> list[Path]:
    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()

    resolved_languages = _resolve_language_specs(languages)
    voice = resolve_voice_asset(
        project_root=project_root,
        voice_name=voice_name,
        ref_audio=ref_audio,
        ref_text=ref_text,
    )

    logger.info("Build workspace: %s", project_root)
    logger.info("Output directory: %s", paths.output)
    logger.info(
        "Voice asset: %s (audio=%s, text=%s)",
        voice.name,
        voice.audio_path,
        voice.text_path,
    )
    logger.info(
        "Languages: %s",
        ", ".join(spec.tag for spec in resolved_languages),
    )
    logger.info(
        "Video settings: %sx%s, padding=%ss, fps=%s",
        slide_width,
        slide_height,
        slide_padding_sec,
        fps,
    )

    _report_progress(progress_callback, "Exporting slide images")
    logger.info("Exporting slide images from %s", pptx_path)
    slide_count = export_slides(
        pptx_path=pptx_path,
        output_dir=paths.slides,
        temp_dir=paths.temp,
        width=slide_width,
        height=slide_height,
    )
    logger.info("Exported %d slide images", slide_count)

    _report_progress(progress_callback, "Generating audio tracks")
    audio_result = generate_audio_tracks(
        pptx_path=pptx_path,
        voice=voice,
        languages=resolved_languages,
        paths=paths,
        model_size=model_size,
        device=device,
        dtype=dtype,
        prompt_cache_path=prompt_cache_path,
        refresh_prompt_cache=refresh_prompt_cache,
        use_prompt_cache=use_prompt_cache,
        progress_callback=progress_callback,
    )
    if audio_result.slide_count != slide_count:
        raise RuntimeError(
            f"Slide export count ({slide_count}) does not match note parsing count ({audio_result.slide_count})."
        )

    _report_progress(progress_callback, "Building videos")
    outputs = build_videos(
        slide_count=audio_result.slide_count,
        languages=audio_result.active_languages,
        paths=paths,
        slide_padding_sec=slide_padding_sec,
        fps=fps,
    )
    for path in outputs:
        logger.info("Created %s", path)

    _report_progress(progress_callback, "Completed")
    return outputs


def _resolve_language_specs(
    languages: Sequence[str | LanguageSpec],
) -> list[LanguageSpec]:
    if languages and all(isinstance(spec, LanguageSpec) for spec in languages):
        return [spec for spec in languages if isinstance(spec, LanguageSpec)]
    return resolve_languages([str(tag) for tag in languages])


def _report_progress(progress_callback: ProgressCallback | None, message: str) -> None:
    if progress_callback is not None:
        progress_callback(message)


def main() -> None:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
    )
    args = build_arg_parser().parse_args()

    build_movie(
        pptx_path=args.pptx,
        project_root=Path.cwd(),
        voice_name=args.voice_name,
        ref_audio=args.ref_audio,
        ref_text=args.ref_text,
        languages=args.languages,
        model_size=args.model_size,
        device=args.device,
        dtype=args.dtype,
        slide_width=args.slide_width,
        slide_height=args.slide_height,
        slide_padding_sec=args.slide_padding_sec,
        fps=args.fps,
        prompt_cache_path=args.prompt_cache,
        refresh_prompt_cache=args.refresh_prompt_cache,
    )


if __name__ == "__main__":
    main()
