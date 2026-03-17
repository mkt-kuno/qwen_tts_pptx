from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path


logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class VoiceAsset:
    name: str
    audio_path: Path
    text_path: Path
    transcript: str


def resolve_voice_asset(
    project_root: Path,
    voice_name: str | None = None,
    ref_audio: Path | None = None,
    ref_text: Path | None = None,
) -> VoiceAsset:
    if voice_name:
        audio_path, text_path = _resolve_named_voice(project_root, voice_name)
    else:
        if ref_audio is None:
            raise ValueError("Provide --voice-name or --ref-audio.")
        audio_path = ref_audio
        text_path = ref_text or ref_audio.with_suffix(".txt")
        if not text_path.exists():
            raise FileNotFoundError(f"Reference transcript not found: {text_path}")

    if not audio_path.exists():
        raise FileNotFoundError(f"Reference audio not found: {audio_path}")
    transcript = _read_text_file(text_path)
    name = voice_name or audio_path.stem
    logger.info(
        "Resolved voice asset '%s' (audio=%s, text=%s)",
        name,
        audio_path,
        text_path,
    )
    return VoiceAsset(
        name=name, audio_path=audio_path, text_path=text_path, transcript=transcript
    )


def _resolve_named_voice(project_root: Path, voice_name: str) -> tuple[Path, Path]:
    candidates = [
        project_root / f"{voice_name}.wav",
        project_root / "assets" / "voices" / f"{voice_name}.wav",
        project_root / "assets" / "voices" / voice_name / f"{voice_name}.wav",
    ]
    for audio_path in candidates:
        text_path = audio_path.with_suffix(".txt")
        if audio_path.exists() and text_path.exists():
            return audio_path, text_path
    searched = ", ".join(str(path) for path in candidates)
    raise FileNotFoundError(f"Voice '{voice_name}' not found. Searched: {searched}")


def _read_text_file(path: Path) -> str:
    encodings = ["utf-8", "utf-8-sig", "cp932"]
    last_error: UnicodeDecodeError | None = None
    for encoding in encodings:
        try:
            return path.read_text(encoding=encoding).strip()
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError(f"Unable to decode transcript file: {path}") from last_error
