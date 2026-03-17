from __future__ import annotations

import logging
import subprocess
import wave
from pathlib import Path

from app.common.languages import LanguageSpec
from app.common.paths import WorkspacePaths


logger = logging.getLogger(__name__)


def build_videos(
    slide_count: int,
    languages: list[LanguageSpec],
    paths: WorkspacePaths,
    slide_padding_sec: float,
    fps: int,
) -> list[Path]:
    if slide_count <= 0:
        raise ValueError("slide_count must be positive")

    durations = compute_slide_durations(
        slide_count=slide_count,
        languages=languages,
        paths=paths,
        slide_padding_sec=slide_padding_sec,
    )
    logger.info("Computed %d slide durations", len(durations))
    outputs: list[Path] = []
    for spec in languages:
        audio_track = paths.temp / f"{spec.directory_name}.wav"
        video_track = paths.temp / f"{spec.directory_name}.m4v"
        output_path = paths.output / spec.output_name
        logger.info("Building %s video assets", spec.tag)
        build_audio_track(
            audio_dir=paths.audio_dir(spec.directory_name),
            durations=durations,
            output_path=audio_track,
        )
        build_slide_video(
            slides_dir=paths.slides,
            durations=durations,
            fps=fps,
            concat_path=paths.temp / f"{spec.directory_name}.ffconcat",
            output_path=video_track,
        )
        mux_video_audio(
            video_path=video_track, audio_path=audio_track, output_path=output_path
        )
        logger.info("Finished %s video: %s", spec.tag, output_path)
        outputs.append(output_path)
    return outputs


def compute_slide_durations(
    slide_count: int,
    languages: list[LanguageSpec],
    paths: WorkspacePaths,
    slide_padding_sec: float,
) -> list[float]:
    durations: list[float] = []
    for slide_number in range(1, slide_count + 1):
        max_duration = 0.0
        for spec in languages:
            wav_path = paths.audio_dir(spec.directory_name) / f"page{slide_number}.wav"
            max_duration = max(max_duration, read_wav_duration(wav_path))
        durations.append(round(max_duration + slide_padding_sec, 3))
    return durations


def read_wav_duration(path: Path) -> float:
    with wave.open(str(path), "rb") as wav_file:
        return wav_file.getnframes() / wav_file.getframerate()


def build_audio_track(
    audio_dir: Path, durations: list[float], output_path: Path
) -> None:
    if not durations:
        raise ValueError("durations must not be empty")

    params: tuple[int, int, int] | None = None
    output_path.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Concatenating audio track %s", output_path)

    with wave.open(str(output_path), "wb") as out_file:
        for slide_number, target_duration in enumerate(durations, start=1):
            wav_path = audio_dir / f"page{slide_number}.wav"
            with wave.open(str(wav_path), "rb") as in_file:
                current_params = (
                    in_file.getnchannels(),
                    in_file.getsampwidth(),
                    in_file.getframerate(),
                )
                if params is None:
                    params = current_params
                    out_file.setnchannels(current_params[0])
                    out_file.setsampwidth(current_params[1])
                    out_file.setframerate(current_params[2])
                elif params != current_params:
                    raise RuntimeError(f"Mismatched WAV format: {wav_path}")

                frames = in_file.readframes(in_file.getnframes())
                out_file.writeframes(frames)
                source_duration = in_file.getnframes() / in_file.getframerate()
                remaining_duration = max(0.0, target_duration - source_duration)
                silence_frames = int(round(remaining_duration * in_file.getframerate()))
                silence_bytes = (
                    b"\x00"
                    * silence_frames
                    * in_file.getnchannels()
                    * in_file.getsampwidth()
                )
                out_file.writeframes(silence_bytes)


def build_slide_video(
    slides_dir: Path,
    durations: list[float],
    fps: int,
    concat_path: Path,
    output_path: Path,
) -> None:
    concat_path.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Rendering slide video %s", output_path)
    lines = ["ffconcat version 1.0"]
    last_path: Path | None = None
    for slide_number, duration in enumerate(durations, start=1):
        slide_path = slides_dir / f"page{slide_number}.png"
        if not slide_path.exists():
            raise FileNotFoundError(f"Slide image not found: {slide_path}")
        escaped = slide_path.resolve().as_posix()
        lines.append(f"file '{escaped}'")
        lines.append(f"duration {duration:.3f}")
        last_path = slide_path
    if last_path is None:
        raise ValueError("No slide images found")
    lines.append(f"file '{last_path.resolve().as_posix()}'")
    concat_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    subprocess.run(
        [
            "ffmpeg",
            "-y",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            str(concat_path),
            "-vf",
            f"fps={fps},format=yuv420p",
            "-an",
            "-c:v",
            "libx264",
            str(output_path),
        ],
        check=True,
    )


def mux_video_audio(video_path: Path, audio_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Muxing %s and %s into %s", video_path, audio_path, output_path)
    subprocess.run(
        [
            "ffmpeg",
            "-y",
            "-i",
            str(video_path),
            "-i",
            str(audio_path),
            "-c:v",
            "copy",
            "-c:a",
            "aac",
            "-b:a",
            "64k",
            str(output_path),
        ],
        check=True,
    )
