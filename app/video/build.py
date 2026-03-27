from __future__ import annotations

import logging
import subprocess
import wave
from pathlib import Path

from app.common.languages import LanguageSpec
from app.common.paths import WorkspacePaths


logger = logging.getLogger(__name__)

LANGUAGE_CODES: dict[str, str] = {
    "EN": "eng",
    "JP": "jpn",
    "ZH": "chi",
    "ES": "spa",
    "IT": "ita",
    "FR": "fre",
}


def build_videos(
    slide_count: int,
    languages: list[LanguageSpec],
    paths: WorkspacePaths,
    slide_padding_sec: float,
    fps: int,
) -> list[Path]:
    if slide_count <= 0:
        raise ValueError("slide_count must be positive")

    outputs: list[Path] = []
    for spec in languages:
        durations = compute_slide_durations_for_language(
            slide_count=slide_count,
            language=spec,
            paths=paths,
            slide_padding_sec=slide_padding_sec,
        )
        logger.info(
            "Computed %d slide durations for %s", len(durations), spec.tag
        )
        audio_track = paths.temp / f"{spec.directory_name}.wav"
        video_track = paths.temp / f"{spec.directory_name}.m4v"
        chapter_path = paths.temp / f"{spec.directory_name}.ffmeta"
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
        write_chapter_metadata(durations=durations, output_path=chapter_path)
        mux_video_audio(
            video_path=video_track,
            audio_path=audio_track,
            output_path=output_path,
            language_code=LANGUAGE_CODES.get(spec.tag, "und"),
            track_title=spec.track_name,
            chapter_path=chapter_path,
        )
        logger.info("Finished %s video: %s", spec.tag, output_path)
        outputs.append(output_path)
    return outputs


def build_multilingual_video(
    slide_count: int,
    languages: list[LanguageSpec],
    paths: WorkspacePaths,
    slide_padding_sec: float,
    fps: int,
    output_name: str = "multilingual.mp4",
) -> Path:
    durations = compute_slide_durations(
        slide_count=slide_count,
        languages=languages,
        paths=paths,
        slide_padding_sec=slide_padding_sec,
    )
    ordered_languages = _order_multilingual_languages(languages)
    output_path = paths.output / output_name
    video_track = paths.temp / "multilingual.m4v"
    concat_path = paths.temp / "multilingual.ffconcat"
    chapter_path = paths.temp / "multilingual.ffmeta"
    build_slide_video(
        slides_dir=paths.slides,
        durations=durations,
        fps=fps,
        concat_path=concat_path,
        output_path=video_track,
    )
    write_chapter_metadata(durations=durations, output_path=chapter_path)

    audio_tracks: list[Path] = []
    for spec in ordered_languages:
        track_path = paths.temp / f"multilingual-{spec.directory_name}.wav"
        build_audio_track(
            audio_dir=paths.audio_dir(spec.directory_name),
            durations=durations,
            output_path=track_path,
        )
        audio_tracks.append(track_path)

    command = ["ffmpeg", "-y", "-i", str(video_track)]
    for audio_track in audio_tracks:
        command.extend(["-i", str(audio_track)])
    command.extend(["-i", str(chapter_path)])

    command.extend(["-map", "0:v:0"])
    for index in range(len(audio_tracks)):
        command.extend(["-map", f"{index + 1}:a:0"])
    command.extend(["-map_metadata", str(len(audio_tracks) + 1)])

    command.extend(["-c:v", "copy", "-c:a", "aac", "-b:a", "64k"])
    for index, spec in enumerate(ordered_languages):
        code = LANGUAGE_CODES.get(spec.tag, "und")
        command.extend([f"-metadata:s:a:{index}", f"language={code}"])
        command.extend([f"-metadata:s:a:{index}", f"title={spec.track_name}"])

    if ordered_languages and ordered_languages[0].tag == "EN":
        command.extend(["-disposition:a:0", "default"])
        for index in range(1, len(ordered_languages)):
            command.extend([f"-disposition:a:{index}", "0"])

    command.append(str(output_path))
    logger.info("Building multilingual video %s", output_path)
    subprocess.run(command, check=True)
    return output_path


def _order_multilingual_languages(languages: list[LanguageSpec]) -> list[LanguageSpec]:
    english = [spec for spec in languages if spec.tag == "EN"]
    others = [spec for spec in languages if spec.tag != "EN"]
    return [*english, *others]


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


def compute_slide_durations_for_language(
    slide_count: int,
    language: LanguageSpec,
    paths: WorkspacePaths,
    slide_padding_sec: float,
) -> list[float]:
    durations: list[float] = []
    for slide_number in range(1, slide_count + 1):
        wav_path = (
            paths.audio_dir(language.directory_name) / f"page{slide_number}.wav"
        )
        duration = read_wav_duration(wav_path)
        durations.append(round(duration + slide_padding_sec, 3))
    return durations


def read_wav_duration(path: Path) -> float:
    with wave.open(str(path), "rb") as wav_file:
        return wav_file.getnframes() / wav_file.getframerate()


def write_chapter_metadata(durations: list[float], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    lines = [";FFMETADATA1"]
    time_ms = 0
    for index, duration in enumerate(durations, start=1):
        start_ms = time_ms
        end_ms = time_ms + int(round(duration * 1000))
        lines.append("")
        lines.append("[CHAPTER]")
        lines.append("TIMEBASE=1/1000")
        lines.append(f"START={start_ms}")
        lines.append(f"END={end_ms}")
        lines.append(f"title=Slide {index}")
        time_ms = end_ms
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


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


def mux_video_audio(
    video_path: Path,
    audio_path: Path,
    output_path: Path,
    language_code: str | None = None,
    track_title: str | None = None,
    chapter_path: Path | None = None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Muxing %s and %s into %s", video_path, audio_path, output_path)
    command = ["ffmpeg", "-y", "-i", str(video_path), "-i", str(audio_path)]
    if chapter_path is not None:
        command.extend(["-i", str(chapter_path)])
    command.extend(["-map", "0:v:0", "-map", "1:a:0"])
    if chapter_path is not None:
        command.extend(["-map_metadata", "2"])
    command.extend(["-c:v", "copy", "-c:a", "aac", "-b:a", "64k"])
    if language_code is not None:
        command.extend(["-metadata:s:a:0", f"language={language_code}"])
    if track_title is not None:
        command.extend(["-metadata:s:a:0", f"title={track_title}"])
    command.append(str(output_path))
    subprocess.run(command, check=True)
