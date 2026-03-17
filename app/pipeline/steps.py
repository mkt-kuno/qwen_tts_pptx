from __future__ import annotations

import hashlib
import json
import logging
import shutil
import sys
import zipfile
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path

from app.common.languages import LanguageSpec, resolve_languages
from app.common.notes import extract_tagged_text, iter_slide_notes
from app.common.paths import WorkspacePaths
from app.models.voices import resolve_voice_asset
from app.slides.export_slides import export_slides
from app.synthesis.qwen import (
    create_voice_clone_prompt,
    detect_device,
    load_model,
    synthesize_to_file,
)
from app.video.build import build_multilingual_video, build_videos


logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ImageExportResult:
    slide_count: int
    slides_dir: Path


@dataclass(frozen=True)
class AudioSynthesisResult:
    slide_count: int
    active_languages: list[LanguageSpec]
    generated_count: int
    cache_hit_count: int


@dataclass(frozen=True)
class VideoBuildResult:
    per_language_outputs: list[Path]
    multilingual_output: Path


def step1_export_images(
    *,
    pptx_path: Path,
    project_root: Path,
    slide_width: int,
    slide_height: int,
) -> ImageExportResult:
    if sys.platform != "win32":
        raise RuntimeError("Step 1 image export requires Windows with PowerPoint COM.")
    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()
    logger.info("Exporting slide images from %s", pptx_path)
    slide_count = export_slides(
        pptx_path=pptx_path,
        output_dir=paths.slides,
        temp_dir=paths.temp,
        width=slide_width,
        height=slide_height,
    )
    logger.info("Exported %d slide images into %s", slide_count, paths.slides)
    return ImageExportResult(slide_count=slide_count, slides_dir=paths.slides)


def step2_synthesize_audio(
    *,
    pptx_path: Path,
    project_root: Path,
    ref_audio: Path,
    ref_text: Path,
    languages: list[str],
    model_size: str,
    device: str | None,
    dtype: str | None,
    force_regenerate: bool = False,
) -> AudioSynthesisResult:
    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()

    voice = resolve_voice_asset(
        project_root=project_root,
        voice_name=None,
        ref_audio=ref_audio,
        ref_text=ref_text,
    )
    language_specs = resolve_languages(languages)
    notes = iter_slide_notes(pptx_path)
    slide_texts: list[tuple[int, dict[str, str]]] = []
    active_tags: set[str] = set()

    for note in notes:
        tagged = {
            spec.tag: extract_tagged_text(note.raw_text, spec.tag)
            for spec in language_specs
        }
        for tag, text in tagged.items():
            if text.strip():
                active_tags.add(tag)
        slide_texts.append((note.slide_number, tagged))

    active_languages = [spec for spec in language_specs if spec.tag in active_tags]
    if not active_languages:
        raise RuntimeError("No matching note tags were found in the presentation.")

    for spec in active_languages:
        paths.audio_dir(spec.directory_name).mkdir(parents=True, exist_ok=True)

    resolved_device = detect_device(device)
    model_result = load_model(
        model_size=model_size, device=resolved_device, dtype=dtype
    )
    model = model_result.model
    resolved_device = model_result.device
    prompt = create_voice_clone_prompt(
        model=model,
        ref_audio=voice.audio_path,
        ref_text=voice.transcript,
    )

    cache_root = paths.work / "cache" / "audio"
    cache_root.mkdir(parents=True, exist_ok=True)
    ref_audio_hash = _sha256_file(voice.audio_path)
    ref_text_hash = _sha256_text(voice.transcript)

    generated_count = 0
    cache_hit_count = 0

    for slide_number, tagged in slide_texts:
        for spec in active_languages:
            text = tagged[spec.tag]
            wav_path = paths.audio_dir(spec.directory_name) / f"page{slide_number}.wav"
            cache_file = cache_root / spec.directory_name / f"page{slide_number}.json"
            cache_key = {
                "slide_number": slide_number,
                "language": spec.tag,
                "script_sha256": _sha256_text(text),
                "model_size": model_size,
                "device": resolved_device,
                "dtype": (dtype or "auto"),
                "ref_audio_sha256": ref_audio_hash,
                "ref_text_sha256": ref_text_hash,
            }

            if not force_regenerate and _is_audio_cache_hit(
                cache_file, wav_path, cache_key
            ):
                cache_hit_count += 1
                continue

            synthesize_to_file(
                model=model,
                prompt=prompt,
                text=text,
                language=spec.qwen_language,
                output_path=wav_path,
            )
            _write_audio_cache(cache_file=cache_file, cache_key=cache_key)
            generated_count += 1

    logger.info(
        "Audio synthesis finished: generated=%d cache_hits=%d",
        generated_count,
        cache_hit_count,
    )
    return AudioSynthesisResult(
        slide_count=len(notes),
        active_languages=active_languages,
        generated_count=generated_count,
        cache_hit_count=cache_hit_count,
    )


def step3_build_videos(
    *,
    project_root: Path,
    languages: list[str],
    slide_padding_sec: float,
    fps: int,
) -> VideoBuildResult:
    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()
    language_specs = resolve_languages(languages)
    slide_count = _count_slides(paths.slides)
    per_language_outputs = build_videos(
        slide_count=slide_count,
        languages=language_specs,
        paths=paths,
        slide_padding_sec=slide_padding_sec,
        fps=fps,
    )
    multilingual_output = build_multilingual_video(
        slide_count=slide_count,
        languages=language_specs,
        paths=paths,
        slide_padding_sec=slide_padding_sec,
        fps=fps,
        output_name="multilingual.mp4",
    )
    return VideoBuildResult(
        per_language_outputs=per_language_outputs,
        multilingual_output=multilingual_output,
    )


def import_slide_zip(*, zip_path: Path, project_root: Path) -> int:
    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()
    _clear_dir_files(paths.slides, "page*.png")
    _extract_zip_safely(zip_path=zip_path, target_dir=paths.slides)
    return _count_slides(paths.slides)


def import_audio_zip(*, zip_path: Path, project_root: Path) -> None:
    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()
    if paths.audio.exists():
        shutil.rmtree(paths.audio)
    paths.audio.mkdir(parents=True, exist_ok=True)
    _extract_zip_safely(zip_path=zip_path, target_dir=paths.audio)


def export_zip_from_dir(*, source_dir: Path, output_zip: Path) -> Path:
    output_zip.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                archive.write(path, arcname=path.relative_to(source_dir))
    return output_zip


def _extract_zip_safely(*, zip_path: Path, target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    target_root = target_dir.resolve()
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            member_path = target_root / member.filename
            resolved = member_path.resolve()
            if not str(resolved).startswith(str(target_root)):
                raise RuntimeError(f"Unsafe ZIP entry path: {member.filename}")
        archive.extractall(target_root)


def _count_slides(slides_dir: Path) -> int:
    slide_count = len(list(slides_dir.glob("page*.png")))
    if slide_count == 0:
        raise RuntimeError(f"No slide images found in {slides_dir}")
    return slide_count


def _clear_dir_files(directory: Path, pattern: str) -> None:
    for path in directory.glob(pattern):
        path.unlink()


def _sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _sha256_file(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as file_obj:
        for chunk in iter(lambda: file_obj.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def _is_audio_cache_hit(
    cache_file: Path, wav_path: Path, expected: dict[str, object]
) -> bool:
    if not cache_file.exists() or not wav_path.exists():
        return False
    try:
        actual = json.loads(cache_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return False
    for key, expected_value in expected.items():
        if actual.get(key) != expected_value:
            return False
    return True


def _write_audio_cache(*, cache_file: Path, cache_key: dict[str, object]) -> None:
    cache_file.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        **cache_key,
        "updated_at": datetime.now(UTC).isoformat(),
        "format": "wav",
    }
    cache_file.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )
