from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class LanguageSpec:
    tag: str
    directory_name: str
    qwen_language: str
    output_name: str


LANGUAGE_SPECS: dict[str, LanguageSpec] = {
    "JA": LanguageSpec(
        tag="JA",
        directory_name="ja",
        qwen_language="Japanese",
        output_name="ja.mp4",
    ),
    "EN": LanguageSpec(
        tag="EN",
        directory_name="en",
        qwen_language="English",
        output_name="en.mp4",
    ),
    "ZH": LanguageSpec(
        tag="ZH",
        directory_name="zh",
        qwen_language="Chinese",
        output_name="zh.mp4",
    ),
}


def resolve_languages(tags: list[str]) -> list[LanguageSpec]:
    resolved: list[LanguageSpec] = []
    for raw_tag in tags:
        tag = raw_tag.upper()
        if tag not in LANGUAGE_SPECS:
            available = ", ".join(sorted(LANGUAGE_SPECS))
            raise ValueError(
                f"Unsupported language tag '{raw_tag}'. Available: {available}"
            )
        resolved.append(LANGUAGE_SPECS[tag])
    return resolved
