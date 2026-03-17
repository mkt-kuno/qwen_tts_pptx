from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class WorkspacePaths:
    root: Path
    work: Path
    output: Path
    slides: Path
    audio: Path
    prompts: Path
    temp: Path

    @classmethod
    def from_root(cls, root: Path) -> "WorkspacePaths":
        return cls(
            root=root,
            work=root / "work",
            output=root / "output",
            slides=root / "work" / "slides",
            audio=root / "work" / "audio",
            prompts=root / "work" / "prompts",
            temp=root / "work" / "temp",
        )

    def ensure_directories(self) -> None:
        for path in [
            self.work,
            self.output,
            self.slides,
            self.audio,
            self.prompts,
            self.temp,
        ]:
            path.mkdir(parents=True, exist_ok=True)

    def audio_dir(self, language_directory: str) -> Path:
        return self.audio / language_directory
