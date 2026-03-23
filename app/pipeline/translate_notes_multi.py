from __future__ import annotations

import argparse
import json
import logging
import re
from pathlib import Path
from urllib import error, request

from app.common.notes import (
    append_tagged_note_blocks,
    extract_tagged_text,
    iter_slide_notes,
    normalize_tagged_note_text,
    write_slide_notes_texts,
)


logger = logging.getLogger(__name__)

TARGET_LANGUAGE_NAMES: dict[str, str] = {
    "EN": "English",
    "ZH": "Chinese",
    "ES": "Spanish",
    "IT": "Italian",
    "FR": "French",
}


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Fill missing EN/ZH/ES/IT/FR note blocks in a PPTX by translating from JP."
        )
    )
    parser.add_argument("pptx", type=Path, help="Input .pptx path")
    parser.add_argument("--model", type=str, default="translategemma:4b")
    parser.add_argument("--host", type=str, default="http://127.0.0.1:11434")
    parser.add_argument("--timeout-sec", type=float, default=60.0)
    return parser


def main() -> None:
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
    )
    args = build_arg_parser().parse_args()

    input_pptx = args.pptx
    output_pptx = input_pptx.with_name(f"{input_pptx.stem}_multi{input_pptx.suffix}")
    logger.info("Reading notes from %s", input_pptx)
    notes = iter_slide_notes(input_pptx)

    missing_jp_slides = [
        note.slide_number
        for note in notes
        if not extract_tagged_text(note.raw_text, "JP").strip()
    ]
    if missing_jp_slides:
        raise RuntimeError(
            "Missing required [JP]... [JP] note block on slides: "
            + ", ".join(str(value) for value in missing_jp_slides)
        )

    notes_text_by_slide: dict[int, str] = {}
    generated_blocks = 0
    for note in notes:
        normalized_raw_text = normalize_tagged_note_text(note.raw_text)
        jp_text = extract_tagged_text(normalized_raw_text, "JP").strip()
        existing_refs = _collect_existing_references(normalized_raw_text)

        missing_tags = [
            tag
            for tag in TARGET_LANGUAGE_NAMES
            if not extract_tagged_text(normalized_raw_text, tag).strip()
        ]
        if not missing_tags:
            notes_text_by_slide[note.slide_number] = normalized_raw_text
            continue

        generated: dict[str, str] = {}
        remaining_tags = list(missing_tags)

        try:
            generated = translate_missing_tags_one_shot(
                source_text=jp_text,
                existing_references=existing_refs,
                missing_tags=missing_tags,
                model=args.model,
                host=args.host,
                timeout_sec=args.timeout_sec,
            )
        except RuntimeError as exc:
            logger.warning(
                "One-shot translation failed on slide %d: %s. Falling back to per-tag.",
                note.slide_number,
                exc,
            )
            generated = {}

        remaining_tags = [tag for tag in missing_tags if tag not in generated]
        if remaining_tags:
            logger.info(
                "Slide %d fallback generation for tags: %s",
                note.slide_number,
                ", ".join(remaining_tags),
            )

        for tag in remaining_tags:
            try:
                translated = translate_text(
                    source_text=jp_text,
                    target_language=TARGET_LANGUAGE_NAMES[tag],
                    existing_references=existing_refs,
                    model=args.model,
                    host=args.host,
                    timeout_sec=args.timeout_sec,
                )
            except RuntimeError as exc:
                raise RuntimeError(
                    f"Translation failed for slide {note.slide_number} tag {tag}: {exc}"
                ) from exc
            generated[tag] = translated

        generated_blocks += len(generated)
        logger.info(
            "Generated %d note blocks for slide %d (%s)",
            len(generated),
            note.slide_number,
            ", ".join(sorted(generated)),
        )

        notes_text_by_slide[note.slide_number] = append_tagged_note_blocks(
            normalized_raw_text, generated
        )

    write_slide_notes_texts(
        pptx_path=input_pptx,
        output_pptx_path=output_pptx,
        notes_text_by_slide=notes_text_by_slide,
    )
    logger.info(
        "Created %s (slides updated=%d generated blocks=%d)",
        output_pptx,
        len(notes_text_by_slide),
        generated_blocks,
    )


def translate_text(
    *,
    source_text: str,
    target_language: str,
    existing_references: dict[str, str],
    model: str,
    host: str,
    timeout_sec: float,
) -> str:
    prompt = _build_translation_prompt(
        source_text=source_text,
        target_language=target_language,
        existing_references=existing_references,
    )
    response = _ollama_generate(
        model=model,
        host=host,
        prompt=prompt,
        timeout_sec=timeout_sec,
    )
    translated = response.strip()
    if not translated:
        raise RuntimeError(
            f"Ollama returned an empty translation for {target_language}."
        )
    return translated


def translate_missing_tags_one_shot(
    *,
    source_text: str,
    existing_references: dict[str, str],
    missing_tags: list[str],
    model: str,
    host: str,
    timeout_sec: float,
) -> dict[str, str]:
    prompt = _build_one_shot_prompt(
        source_text=source_text,
        existing_references=existing_references,
        missing_tags=missing_tags,
    )
    response = _ollama_generate(
        model=model,
        host=host,
        prompt=prompt,
        timeout_sec=timeout_sec,
    )
    parsed = _parse_one_shot_response(
        response_text=response, required_tags=missing_tags
    )
    return parsed


def _build_translation_prompt(
    *,
    source_text: str,
    target_language: str,
    existing_references: dict[str, str],
) -> str:
    return (
        "You are a translation engine.\n"
        f"Translate the Japanese text directly into {target_language}.\n"
        "Rules:\n"
        "- Output only the translated text.\n"
        "- Do not add labels, explanations, or quotation marks.\n"
        "- Do not translate from any reference language text.\n"
        "- Use reference translations only for terminology consistency.\n"
        "- Preserve original line breaks and paragraph boundaries.\n"
        "- Keep numbers, symbols, and markdown-like tokens unchanged where possible.\n\n"
        "Existing reference translations from the original note (non-JP only):\n"
        f"{_format_existing_references(existing_references)}\n\n"
        "Japanese source:\n"
        f"{source_text}"
    )


def _build_one_shot_prompt(
    *,
    source_text: str,
    existing_references: dict[str, str],
    missing_tags: list[str],
) -> str:
    missing_list = ", ".join(missing_tags)
    return (
        "You are a translation engine.\n"
        "Translate the Japanese source text directly into all requested languages.\n"
        "Use reference translations only for terminology consistency and style alignment.\n"
        "Never translate from the reference language texts.\n\n"
        "Requirements:\n"
        f"- Generate only these tags: {missing_list}.\n"
        "- Return a strict JSON object only.\n"
        "- JSON keys must be exactly the requested tags and values must be translated strings.\n"
        "- Do not include markdown, code fences, comments, or extra keys.\n"
        "- Preserve original line breaks and paragraph boundaries where possible.\n"
        "- Keep numbers, symbols, and markdown-like tokens unchanged where possible.\n\n"
        "Existing reference translations from the original note (non-JP only):\n"
        f"{_format_existing_references(existing_references)}\n\n"
        "Japanese source:\n"
        f"{source_text}"
    )


def _collect_existing_references(raw_text: str) -> dict[str, str]:
    references: dict[str, str] = {}
    for tag in TARGET_LANGUAGE_NAMES:
        text = extract_tagged_text(raw_text, tag).strip()
        if text:
            references[tag] = text
    return references


def _format_existing_references(existing_references: dict[str, str]) -> str:
    if not existing_references:
        return "(none)"
    lines = []
    for tag in TARGET_LANGUAGE_NAMES:
        value = existing_references.get(tag)
        if value:
            lines.append(f"[{tag}]\n{value}\n[{tag}]")
    return "\n\n".join(lines)


def _parse_one_shot_response(
    *, response_text: str, required_tags: list[str]
) -> dict[str, str]:
    parsed = _extract_json_object(response_text)
    generated: dict[str, str] = {}
    invalid_tags: list[str] = []
    for tag in required_tags:
        value = parsed.get(tag)
        if isinstance(value, str) and value.strip():
            generated[tag] = value.strip()
        else:
            invalid_tags.append(tag)

    extra_keys = sorted(key for key in parsed if key not in required_tags)
    if extra_keys:
        logger.warning(
            "One-shot response included unexpected keys; ignoring extras: %s",
            ", ".join(extra_keys),
        )

    if invalid_tags:
        logger.warning(
            "One-shot response missing/invalid tags: %s",
            ", ".join(invalid_tags),
        )
    return generated


def _extract_json_object(response_text: str) -> dict[str, object]:
    candidates: list[str] = [response_text.strip()]

    fence_pattern = re.compile(r"```(?:json)?\s*(.*?)```", re.IGNORECASE | re.DOTALL)
    for match in fence_pattern.finditer(response_text):
        fenced = match.group(1).strip()
        if fenced:
            candidates.append(fenced)

    for candidate in candidates:
        parsed = _try_load_json_object(candidate)
        if parsed is not None:
            return parsed
        for snippet in _iter_braced_json_candidates(candidate):
            parsed = _try_load_json_object(snippet)
            if parsed is not None:
                return parsed

    raise RuntimeError("Could not extract a valid JSON object from model response.")


def _try_load_json_object(text: str) -> dict[str, object] | None:
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        return None
    if isinstance(parsed, dict):
        return parsed
    return None


def _iter_braced_json_candidates(text: str) -> list[str]:
    snippets: list[str] = []
    depth = 0
    start_index: int | None = None
    in_string = False
    escaped = False

    for index, char in enumerate(text):
        if in_string:
            if escaped:
                escaped = False
                continue
            if char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue

        if char == '"':
            in_string = True
            continue

        if char == "{":
            if depth == 0:
                start_index = index
            depth += 1
            continue

        if char == "}" and depth > 0:
            depth -= 1
            if depth == 0 and start_index is not None:
                snippets.append(text[start_index : index + 1])
                start_index = None

    return snippets


def _ollama_generate(*, model: str, host: str, prompt: str, timeout_sec: float) -> str:
    endpoint = host.rstrip("/") + "/api/generate"
    payload = json.dumps(
        {"model": model, "prompt": prompt, "stream": False}, ensure_ascii=False
    ).encode("utf-8")
    http_request = request.Request(
        endpoint,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with request.urlopen(http_request, timeout=timeout_sec) as http_response:
            body = http_response.read().decode("utf-8")
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace") if exc.fp else ""
        raise RuntimeError(
            f"Ollama request failed with HTTP {exc.code} for {endpoint}: {details}"
        ) from exc
    except error.URLError as exc:
        raise RuntimeError(
            f"Failed to connect to Ollama at {endpoint}: {exc.reason}"
        ) from exc
    except TimeoutError as exc:
        raise RuntimeError(
            f"Ollama request timed out after {timeout_sec} seconds for {endpoint}."
        ) from exc

    try:
        parsed = json.loads(body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Ollama returned invalid JSON: {body}") from exc

    if parsed.get("error"):
        raise RuntimeError(f"Ollama error: {parsed['error']}")

    response = parsed.get("response")
    if not isinstance(response, str):
        raise RuntimeError(f"Ollama response missing 'response' text: {parsed}")
    return response


if __name__ == "__main__":
    main()
