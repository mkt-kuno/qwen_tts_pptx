from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class LanguageSpec:
    tag: str
    directory_name: str
    qwen_language: str
    output_name: str


LANGUAGE_SPECS: dict[str, LanguageSpec] = {
    "EN": LanguageSpec(
        tag="EN",
        directory_name="en",
        qwen_language="English",
        output_name="en.mp4",
    ),
    "JP": LanguageSpec(
        tag="JP",
        directory_name="jp",
        qwen_language="Japanese",
        output_name="jp.mp4",
    ),
    "ZH": LanguageSpec(
        tag="ZH",
        directory_name="zh",
        qwen_language="Chinese",
        output_name="zh.mp4",
    ),
    "ES": LanguageSpec(
        tag="ES",
        directory_name="es",
        qwen_language="Spanish",
        output_name="es.mp4",
    ),
    "IT": LanguageSpec(
        tag="IT",
        directory_name="it",
        qwen_language="Italian",
        output_name="it.mp4",
    ),
    "FR": LanguageSpec(
        tag="FR",
        directory_name="fr",
        qwen_language="French",
        output_name="fr.mp4",
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
