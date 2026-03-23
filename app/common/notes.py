from __future__ import annotations

import posixpath
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
PPT_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
TAG_PATTERN = re.compile(
    r"\[(?P<tag>EN|JP|ZH|ES|IT|FR)\](.*?)\[(?P=tag)\]",
    re.DOTALL | re.IGNORECASE,
)
SKIPPED_PLACEHOLDER_TYPES = {"dt", "ftr", "hdr", "sldImg", "sldNum"}
NOTES_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
)
XML_NS = "http://www.w3.org/XML/1998/namespace"


@dataclass(frozen=True)
class SlideNotes:
    slide_number: int
    raw_text: str


def extract_tagged_text(text: str, tag: str) -> str:
    if not text:
        return ""
    parts: list[str] = []
    for match in TAG_PATTERN.finditer(text):
        if match.group("tag").lower() == tag.lower():
            value = match.group(2).strip()
            if value:
                parts.append(value)
    return "\n".join(parts).strip()


def iter_slide_notes(pptx_path: Path) -> list[SlideNotes]:
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")

    with zipfile.ZipFile(pptx_path) as archive:
        slide_paths = _iter_slide_paths(archive)
        notes: list[SlideNotes] = []
        for slide_number, slide_path in enumerate(slide_paths, start=1):
            notes_path = _find_notes_slide_path(archive, slide_path)
            raw_text = _extract_notes_text(archive, notes_path) if notes_path else ""
            notes.append(SlideNotes(slide_number=slide_number, raw_text=raw_text))
        return notes


def append_tagged_note_blocks(raw_text: str, blocks: dict[str, str]) -> str:
    updated = normalize_tagged_note_text(raw_text)
    for tag, value in blocks.items():
        text = _normalize_newlines(value).strip()
        if not text:
            continue
        block = _format_tagged_block(tag=tag.upper(), text=text)
        if not updated:
            updated = block
        else:
            updated = f"{updated}\n\n{block}"
    return updated


def normalize_tagged_note_text(raw_text: str) -> str:
    normalized_source = _normalize_newlines(raw_text)
    normalized_parts: list[str] = []
    position = 0

    for match in TAG_PATTERN.finditer(normalized_source):
        leading_text = normalized_source[position : match.start()].strip()
        if leading_text:
            normalized_parts.append(leading_text)

        tag = match.group("tag").upper()
        content = _normalize_newlines(match.group(2)).strip()
        normalized_parts.append(_format_tagged_block(tag=tag, text=content))
        position = match.end()

    trailing_text = normalized_source[position:].strip()
    if trailing_text:
        normalized_parts.append(trailing_text)

    return "\n\n".join(normalized_parts).strip()


def write_slide_notes_texts(
    *,
    pptx_path: Path,
    output_pptx_path: Path,
    notes_text_by_slide: dict[int, str],
) -> Path:
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")

    if not notes_text_by_slide:
        output_pptx_path.parent.mkdir(parents=True, exist_ok=True)
        with (
            zipfile.ZipFile(pptx_path) as source_zip,
            zipfile.ZipFile(output_pptx_path, "w") as output_zip,
        ):
            for entry in source_zip.infolist():
                output_zip.writestr(entry, source_zip.read(entry.filename))
        return output_pptx_path

    with zipfile.ZipFile(pptx_path) as archive:
        slide_paths = _iter_slide_paths(archive)
        notes_parts_by_slide: dict[int, str] = {}
        for slide_number, slide_path in enumerate(slide_paths, start=1):
            notes_path = _find_notes_slide_path(archive, slide_path)
            if notes_path:
                notes_parts_by_slide[slide_number] = notes_path

        unknown_slides = sorted(
            slide_number
            for slide_number in notes_text_by_slide
            if slide_number not in notes_parts_by_slide
        )
        if unknown_slides:
            raise RuntimeError(
                "Cannot update notes for slides without notes parts: "
                + ", ".join(str(value) for value in unknown_slides)
            )

        rewritten_parts: dict[str, bytes] = {}
        for slide_number, updated_text in notes_text_by_slide.items():
            notes_part = notes_parts_by_slide[slide_number]
            rewritten_parts[notes_part] = _rewrite_notes_part_text(
                notes_xml=archive.read(notes_part), updated_text=updated_text
            )

        output_pptx_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_pptx_path, "w") as output_zip:
            for entry in archive.infolist():
                payload = rewritten_parts.get(entry.filename)
                if payload is None:
                    payload = archive.read(entry.filename)
                output_zip.writestr(entry, payload)

    return output_pptx_path


def _iter_slide_paths(archive: zipfile.ZipFile) -> list[str]:
    presentation_xml = ET.fromstring(archive.read("ppt/presentation.xml"))
    relationships_xml = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))

    relationship_map = {
        node.attrib["Id"]: _normalize_part_path(
            "ppt/presentation.xml", node.attrib["Target"]
        )
        for node in relationships_xml.findall("rel:Relationship", REL_NS)
    }

    slide_paths: list[str] = []
    for node in presentation_xml.findall("./p:sldIdLst/p:sldId", PPT_NS):
        rel_id = node.attrib[f"{{{PPT_NS['r']}}}id"]
        slide_paths.append(relationship_map[rel_id])
    return slide_paths


def _find_notes_slide_path(archive: zipfile.ZipFile, slide_path: str) -> str | None:
    rels_path = _part_rels_path(slide_path)
    if rels_path not in archive.namelist():
        return None

    relationships_xml = ET.fromstring(archive.read(rels_path))
    for node in relationships_xml.findall("rel:Relationship", REL_NS):
        if node.attrib.get("Type") == NOTES_REL_TYPE:
            return _normalize_part_path(slide_path, node.attrib["Target"])
    return None


def _extract_notes_text(archive: zipfile.ZipFile, notes_path: str) -> str:
    root = ET.fromstring(archive.read(notes_path))
    lines: list[str] = []
    for shape in root.findall(".//p:sp", PPT_NS):
        placeholder = shape.find("./p:nvSpPr/p:nvPr/p:ph", PPT_NS)
        placeholder_type = (
            placeholder.attrib.get("type") if placeholder is not None else None
        )
        if placeholder_type in SKIPPED_PLACEHOLDER_TYPES:
            continue

        for paragraph in shape.findall(".//a:p", PPT_NS):
            text = "".join(
                node.text or "" for node in paragraph.findall(".//a:t", PPT_NS)
            ).strip()
            if text:
                lines.append(text)
    return "\n".join(lines).strip()


def _rewrite_notes_part_text(*, notes_xml: bytes, updated_text: str) -> bytes:
    root = ET.fromstring(notes_xml)
    target_shape = _find_editable_notes_shape(root)
    if target_shape is None:
        raise RuntimeError("No editable text shape found in notes slide.")

    text_body = target_shape.find("./p:txBody", PPT_NS)
    if text_body is None:
        raise RuntimeError("Notes slide text shape does not contain p:txBody.")

    for paragraph in list(text_body.findall("./a:p", PPT_NS)):
        text_body.remove(paragraph)

    lines = updated_text.splitlines() if updated_text else [""]
    for line in lines:
        paragraph = ET.SubElement(text_body, _a_tag("p"))
        if line:
            run = ET.SubElement(paragraph, _a_tag("r"))
            text_node = ET.SubElement(run, _a_tag("t"))
            text_node.text = line
            text_node.set(f"{{{XML_NS}}}space", "preserve")
        ET.SubElement(paragraph, _a_tag("endParaRPr"))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _find_editable_notes_shape(root: ET.Element) -> ET.Element | None:
    for shape in root.findall(".//p:sp", PPT_NS):
        placeholder = shape.find("./p:nvSpPr/p:nvPr/p:ph", PPT_NS)
        placeholder_type = (
            placeholder.attrib.get("type") if placeholder is not None else None
        )
        if placeholder_type in SKIPPED_PLACEHOLDER_TYPES:
            continue
        if shape.find("./p:txBody", PPT_NS) is not None:
            return shape
    return None


def _a_tag(local_name: str) -> str:
    return f"{{{PPT_NS['a']}}}{local_name}"


def _format_tagged_block(*, tag: str, text: str) -> str:
    if text:
        return f"[{tag}]\n{text}\n[{tag}]"
    return f"[{tag}]\n[{tag}]"


def _normalize_newlines(value: str) -> str:
    return value.replace("\r\n", "\n").replace("\r", "\n")


def _normalize_part_path(base_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _part_rels_path(part_path: str) -> str:
    directory = posixpath.dirname(part_path)
    filename = posixpath.basename(part_path)
    return posixpath.join(directory, "_rels", f"{filename}.rels")
