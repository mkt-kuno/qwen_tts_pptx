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
    r"\[(?P<tag>JA|EN|ZH)\](.*?)\[(?P=tag)\]", re.DOTALL | re.IGNORECASE
)
SKIPPED_PLACEHOLDER_TYPES = {"dt", "ftr", "hdr", "sldImg", "sldNum"}
NOTES_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
)


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


def _normalize_part_path(base_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _part_rels_path(part_path: str) -> str:
    directory = posixpath.dirname(part_path)
    filename = posixpath.basename(part_path)
    return posixpath.join(directory, "_rels", f"{filename}.rels")
