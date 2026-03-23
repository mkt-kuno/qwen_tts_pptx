from __future__ import annotations

import audioop
import hashlib
import json
import logging
import math
import shutil
import sys
import wave
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from app.common.languages import LanguageSpec, resolve_languages
from app.common.notes import extract_tagged_text, iter_slide_notes
from app.common.paths import WorkspacePaths
from app.models.voices import resolve_voice_asset
from app.slides.export_slides import export_slides
from app.synthesis.qwen import (
    SAFE_VOICE_CLONE_GENERATION_PARAMS,
    VoiceCloneGenerationParams,
    create_voice_clone_prompt,
    detect_device,
    load_model,
    resolve_voice_clone_language,
    synthesize_batch_to_files,
    synthesize_to_file,
)
from app.video.build import build_multilingual_video, build_videos


logger = logging.getLogger(__name__)

ANOMALOUS_AUDIO_RETRY_LANGUAGE_TAGS = frozenset({"JP", "EN", "ZH"})
ANOMALOUS_AUDIO_LONG_MIN_DURATION_SEC = 12.0
ANOMALOUS_AUDIO_MAX_SECONDS_PER_CHAR = 0.45
ANOMALOUS_AUDIO_MIN_SECONDS_PER_CHAR = 0.02
ANOMALOUS_AUDIO_SHORT_MIN_DURATION_SEC = 0.30
ANOMALOUS_AUDIO_SILENCE_PEAK_RATIO = 0.001
ANOMALOUS_AUDIO_SILENCE_RMS_RATIO = 0.0005
ANOMALOUS_AUDIO_RETRY_LIMIT = 1


@dataclass(frozen=True)
class _AudioMemoryCacheEntry:
    wav_path: Path
    wav_sha256: str


@dataclass(frozen=True)
class _PendingAudioGeneration:
    slide_number: int
    language_tag: str
    qwen_language: str
    text: str
    wav_path: Path
    cache_key: str


@dataclass(frozen=True)
class _AudioMetrics:
    duration_sec: float
    seconds_per_char: float
    peak_ratio: float
    rms_ratio: float


@dataclass(frozen=True)
class _AudioAnomaly:
    kind: str
    metrics: _AudioMetrics


_AUDIO_MEMORY_CACHE: dict[str, _AudioMemoryCacheEntry] = {}


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
class Step2ProgressUpdate:
    completed_units: int
    total_units: int
    slide_number: int
    language_tag: str


@dataclass(frozen=True)
class VideoBuildResult:
    per_language_outputs: list[Path]
    multilingual_output: Path | None


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
    progress_callback: Callable[[Step2ProgressUpdate], None] | None = None,
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
    active_tag_set = {spec.tag for spec in active_languages}

    total_units = len(slide_texts) * len(language_specs)
    completed_units = 0

    for spec in active_languages:
        paths.audio_dir(spec.directory_name).mkdir(parents=True, exist_ok=True)

    resolved_device = detect_device(device)
    model_result = load_model(
        model_size=model_size, device=resolved_device, dtype=dtype
    )
    model = model_result.model
    resolved_device = model_result.device
    resolved_dtype = model_result.dtype_name
    generation_params = SAFE_VOICE_CLONE_GENERATION_PARAMS
    logger.info(
        "Using conservative generation params: temperature=%.2f top_p=%.2f top_k=%d repetition_penalty=%.2f dtype=%s",
        generation_params.temperature,
        generation_params.top_p,
        generation_params.top_k,
        generation_params.repetition_penalty,
        resolved_dtype,
    )
    prompt = create_voice_clone_prompt(
        model=model,
        ref_audio=voice.audio_path,
        ref_text=voice.transcript,
    )

    ref_audio_hash = _sha256_file(voice.audio_path)
    ref_text_hash = _sha256_text(voice.transcript)
    language_assignments = {
        spec.tag: resolve_voice_clone_language(model, spec.qwen_language)
        for spec in active_languages
    }
    retry_for_anomalies = ANOMALOUS_AUDIO_RETRY_LANGUAGE_TAGS.issubset(
        {spec.tag for spec in active_languages}
    )

    generated_count = 0
    cache_hit_count = 0
    retry_count = 0

    for slide_number, tagged in slide_texts:
        pending: list[_PendingAudioGeneration] = []
        for spec in language_specs:
            _emit_step2_progress(
                callback=progress_callback,
                completed_units=completed_units,
                total_units=total_units,
                slide_number=slide_number,
                language_tag=spec.tag,
            )
            if spec.tag not in active_tag_set:
                completed_units += 1
                continue

            text = tagged[spec.tag]
            qwen_language = language_assignments[spec.tag]
            wav_path = paths.audio_dir(spec.directory_name) / f"page{slide_number}.wav"
            cache_key = _build_audio_memory_cache_key(
                text=text,
                language_tag=spec.tag,
                qwen_language=qwen_language,
                model_size=model_size,
                device=resolved_device,
                dtype_name=resolved_dtype,
                ref_audio_sha256=ref_audio_hash,
                ref_text_sha256=ref_text_hash,
                generation_params=generation_params,
            )

            if not force_regenerate and _restore_audio_from_memory_cache(
                cache_key=cache_key,
                wav_path=wav_path,
            ):
                cache_hit_count += 1
                completed_units += 1
                continue

            pending.append(
                _PendingAudioGeneration(
                    slide_number=slide_number,
                    language_tag=spec.tag,
                    qwen_language=qwen_language,
                    text=text,
                    wav_path=wav_path,
                    cache_key=cache_key,
                )
            )

        if not pending:
            continue

        synthesize_batch_to_files(
            model=model,
            prompt=prompt,
            texts=[item.text for item in pending],
            languages=[item.qwen_language for item in pending],
            output_paths=[item.wav_path for item in pending],
            generation_params=generation_params,
        )

        generated_count += len(pending)

        if retry_for_anomalies:
            for item in pending:
                retry_count += _retry_if_anomalous_audio(
                    model=model,
                    prompt=prompt,
                    slide_number=item.slide_number,
                    text=item.text,
                    language=item.qwen_language,
                    output_path=item.wav_path,
                    generation_params=generation_params,
                )

        for item in pending:
            anomaly = (
                _detect_audio_anomaly(text=item.text, wav_path=item.wav_path)
                if retry_for_anomalies
                else None
            )
            if anomaly is not None:
                logger.warning(
                    "Skipping memory cache for anomalous audio (slide=%d language=%s reason=%s).",
                    item.slide_number,
                    item.language_tag,
                    anomaly.kind,
                )
                completed_units += 1
                continue
            _store_audio_memory_cache(cache_key=item.cache_key, wav_path=item.wav_path)
            completed_units += 1

    logger.info(
        "Audio synthesis finished: generated=%d cache_hits=%d retries=%d",
        generated_count,
        cache_hit_count,
        retry_count,
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
    output_language_tags: list[str] | None = None,
    include_multilingual: bool = True,
) -> VideoBuildResult:
    if not include_multilingual and output_language_tags is not None:
        if len(output_language_tags) == 0:
            raise ValueError("Select at least one video output.")

    paths = WorkspacePaths.from_root(project_root)
    paths.ensure_directories()
    language_specs = resolve_languages(languages)
    selected_language_specs = _resolve_output_language_specs(
        requested_tags=output_language_tags,
        available_languages=language_specs,
    )

    if not include_multilingual and not selected_language_specs:
        raise ValueError("Select at least one video output.")

    slide_count = _count_slides(paths.slides)
    per_language_outputs: list[Path] = []
    if selected_language_specs:
        per_language_outputs = build_videos(
            slide_count=slide_count,
            languages=selected_language_specs,
            paths=paths,
            slide_padding_sec=slide_padding_sec,
            fps=fps,
        )

    multilingual_output: Path | None = None
    if include_multilingual:
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


def _build_audio_memory_cache_key(
    *,
    text: str,
    language_tag: str,
    qwen_language: str,
    model_size: str,
    device: str,
    dtype_name: str,
    ref_audio_sha256: str,
    ref_text_sha256: str,
    generation_params: VoiceCloneGenerationParams,
) -> str:
    payload = {
        "text_sha256": _sha256_text(text),
        "language": language_tag,
        "qwen_language": qwen_language,
        "temperature": generation_params.temperature,
        "top_p": generation_params.top_p,
        "top_k": generation_params.top_k,
        "repetition_penalty": generation_params.repetition_penalty,
        "do_sample": generation_params.do_sample,
        "dtype": dtype_name,
        "model_size": model_size,
        "device": device,
        "ref_audio_sha256": ref_audio_sha256,
        "ref_text_sha256": ref_text_sha256,
    }
    return hashlib.sha256(
        json.dumps(payload, ensure_ascii=True, sort_keys=True).encode("utf-8")
    ).hexdigest()


def _restore_audio_from_memory_cache(*, cache_key: str, wav_path: Path) -> bool:
    entry = _AUDIO_MEMORY_CACHE.get(cache_key)
    if entry is None:
        return False
    if not entry.wav_path.exists():
        _AUDIO_MEMORY_CACHE.pop(cache_key, None)
        return False

    current_hash = _sha256_file(entry.wav_path)
    if current_hash != entry.wav_sha256:
        _AUDIO_MEMORY_CACHE.pop(cache_key, None)
        return False

    if entry.wav_path.resolve() != wav_path.resolve():
        wav_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(entry.wav_path, wav_path)
    return True


def _store_audio_memory_cache(*, cache_key: str, wav_path: Path) -> None:
    _AUDIO_MEMORY_CACHE[cache_key] = _AudioMemoryCacheEntry(
        wav_path=wav_path,
        wav_sha256=_sha256_file(wav_path),
    )


def _retry_if_anomalous_audio(
    *,
    model,
    prompt,
    slide_number: int,
    text: str,
    language: str,
    output_path: Path,
    generation_params: VoiceCloneGenerationParams,
) -> int:
    anomaly = _detect_audio_anomaly(text=text, wav_path=output_path)
    if anomaly is None:
        return 0

    retry_count = 0
    for attempt in range(1, ANOMALOUS_AUDIO_RETRY_LIMIT + 1):
        logger.warning(
            "Anomalous audio detected (slide=%d language=%s reason=%s duration=%.2fs sec_per_char=%.3f peak=%.4f rms=%.4f). Retrying (%d/%d).",
            slide_number,
            language,
            anomaly.kind,
            anomaly.metrics.duration_sec,
            anomaly.metrics.seconds_per_char,
            anomaly.metrics.peak_ratio,
            anomaly.metrics.rms_ratio,
            attempt,
            ANOMALOUS_AUDIO_RETRY_LIMIT,
        )
        synthesize_to_file(
            model=model,
            prompt=prompt,
            text=text,
            language=language,
            output_path=output_path,
            generation_params=generation_params,
        )
        retry_count += 1
        anomaly = _detect_audio_anomaly(text=text, wav_path=output_path)
        if anomaly is None:
            return retry_count

    logger.warning(
        "Audio still appears anomalous after retries (slide=%d language=%s reason=%s duration=%.2fs sec_per_char=%.3f peak=%.4f rms=%.4f).",
        slide_number,
        language,
        anomaly.kind,
        anomaly.metrics.duration_sec,
        anomaly.metrics.seconds_per_char,
        anomaly.metrics.peak_ratio,
        anomaly.metrics.rms_ratio,
    )
    return retry_count


def _detect_audio_anomaly(*, text: str, wav_path: Path) -> _AudioAnomaly | None:
    text_length = _normalized_text_length(text)
    if text_length == 0:
        return None

    try:
        metrics = _read_audio_metrics(wav_path=wav_path, text_length=text_length)
    except (FileNotFoundError, OSError, EOFError, wave.Error):
        return _AudioAnomaly(
            kind="invalid_audio",
            metrics=_AudioMetrics(
                duration_sec=0.0,
                seconds_per_char=0.0,
                peak_ratio=0.0,
                rms_ratio=0.0,
            ),
        )

    if metrics.duration_sec <= 0.0:
        return _AudioAnomaly(kind="silent_audio", metrics=metrics)
    if (
        metrics.peak_ratio <= ANOMALOUS_AUDIO_SILENCE_PEAK_RATIO
        and metrics.rms_ratio <= ANOMALOUS_AUDIO_SILENCE_RMS_RATIO
    ):
        return _AudioAnomaly(kind="silent_audio", metrics=metrics)
    if (
        metrics.duration_sec >= ANOMALOUS_AUDIO_LONG_MIN_DURATION_SEC
        and metrics.seconds_per_char > ANOMALOUS_AUDIO_MAX_SECONDS_PER_CHAR
    ):
        return _AudioAnomaly(kind="too_long", metrics=metrics)
    if (
        metrics.duration_sec < ANOMALOUS_AUDIO_SHORT_MIN_DURATION_SEC
        and text_length >= 4
    ):
        return _AudioAnomaly(kind="too_short", metrics=metrics)
    if metrics.seconds_per_char < ANOMALOUS_AUDIO_MIN_SECONDS_PER_CHAR:
        return _AudioAnomaly(kind="too_short", metrics=metrics)

    return None


def _read_audio_metrics(*, wav_path: Path, text_length: int) -> _AudioMetrics:
    with wave.open(str(wav_path), "rb") as wav_file:
        channels = max(1, wav_file.getnchannels())
        sample_width = wav_file.getsampwidth()
        sample_rate = wav_file.getframerate()
        frame_count = wav_file.getnframes()

        if sample_width <= 0 or sample_rate <= 0 or frame_count <= 0:
            return _AudioMetrics(
                duration_sec=0.0,
                seconds_per_char=0.0,
                peak_ratio=0.0,
                rms_ratio=0.0,
            )

        duration_sec = frame_count / sample_rate
        bytes_per_frame = sample_width * channels
        remaining_frames = frame_count
        total_samples = 0
        sum_rms_square = 0.0
        peak_abs = 0

        while remaining_frames > 0:
            raw = wav_file.readframes(min(remaining_frames, 8192))
            if not raw:
                break
            frames_read = len(raw) // bytes_per_frame
            remaining_frames -= frames_read

            sample_count = len(raw) // sample_width
            if sample_count == 0:
                continue
            chunk_rms = audioop.rms(raw, sample_width)
            chunk_peak = audioop.max(raw, sample_width)
            sum_rms_square += float(chunk_rms * chunk_rms) * float(sample_count)
            total_samples += sample_count
            peak_abs = max(peak_abs, chunk_peak)

        if total_samples == 0:
            return _AudioMetrics(
                duration_sec=duration_sec,
                seconds_per_char=duration_sec / text_length,
                peak_ratio=0.0,
                rms_ratio=0.0,
            )

        full_scale = float((1 << (8 * sample_width - 1)) - 1)
        rms_abs = math.sqrt(sum_rms_square / float(total_samples))
        return _AudioMetrics(
            duration_sec=duration_sec,
            seconds_per_char=duration_sec / text_length,
            peak_ratio=(peak_abs / full_scale) if full_scale > 0 else 0.0,
            rms_ratio=(rms_abs / full_scale) if full_scale > 0 else 0.0,
        )


def _normalized_text_length(text: str) -> int:
    return len("".join(text.split()))


def _emit_step2_progress(
    *,
    callback: Callable[[Step2ProgressUpdate], None] | None,
    completed_units: int,
    total_units: int,
    slide_number: int,
    language_tag: str,
) -> None:
    if callback is None:
        return
    callback(
        Step2ProgressUpdate(
            completed_units=completed_units,
            total_units=total_units,
            slide_number=slide_number,
            language_tag=language_tag,
        )
    )


def _resolve_output_language_specs(
    *,
    requested_tags: list[str] | None,
    available_languages: list[LanguageSpec],
) -> list[LanguageSpec]:
    if requested_tags is None:
        return available_languages

    available_by_tag = {spec.tag: spec for spec in available_languages}
    selected_specs: list[LanguageSpec] = []
    for raw_tag in requested_tags:
        tag = raw_tag.upper()
        spec = available_by_tag.get(tag)
        if spec is None:
            available = ", ".join(spec.tag for spec in available_languages)
            raise ValueError(
                f"Output language '{raw_tag}' is not included in selected languages ({available})."
            )
        selected_specs.append(spec)
    return selected_specs
