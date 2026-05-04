#!/usr/bin/env python3
"""Remove Word navigation outline levels while preserving heading-like styling.

This tool edits .docx OOXML directly. In ALL mode, paragraphs that use styles
with outline levels are moved to cloned "No Nav" styles. The original heading
styles remain available in Word so real headings can be manually restored.
"""

from __future__ import annotations

import argparse
import copy
import os
import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
W_ATTR = f"{{{W_NS}}}"

COMMON_NAMESPACES = {
    "w": W_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
}


for prefix, uri in COMMON_NAMESPACES.items():
    ET.register_namespace(prefix, uri)


@dataclass(frozen=True)
class CleanReport:
    input_path: Path
    output_path: Path
    changed_paragraphs: int
    cloned_styles: int
    direct_outline_removed: int

    def to_text(self) -> str:
        return "\n".join(
            [
                f"input: {self.input_path}",
                f"output: {self.output_path}",
                f"changed_paragraphs: {self.changed_paragraphs}",
                f"cloned_styles: {self.cloned_styles}",
                f"direct_outline_removed: {self.direct_outline_removed}",
            ]
        )


class DocxCleanerError(RuntimeError):
    """Raised for user-facing cleaner failures."""


def default_output_path(input_path: Path | str) -> Path:
    """Return the default cleaned output path next to the input document."""

    source = Path(input_path).expanduser().resolve()
    return source.with_name(f"{source.stem}.cleaned{source.suffix}")


def clean_docx_all(
    input_path: Path | str, output_path: Path | str | None = None, overwrite: bool = False
) -> CleanReport:
    """Clean a .docx in ALL mode and write a new .docx.

    The source file is never modified.
    """

    source = Path(input_path).expanduser().resolve()
    if output_path is None:
        target = default_output_path(source)
    else:
        target = Path(output_path).expanduser().resolve()

    _validate_paths(source, target, overwrite=overwrite)

    try:
        with zipfile.ZipFile(source, "r") as zin:
            names = set(zin.namelist())
            _require_docx_parts(names)
            styles_xml = zin.read("word/styles.xml")
            document_xml = zin.read("word/document.xml")
    except zipfile.BadZipFile as exc:
        raise DocxCleanerError(f"not a valid .docx zip file: {source}") from exc
    except KeyError as exc:
        raise DocxCleanerError(f"missing required .docx part: {exc}") from exc

    try:
        styles_root = ET.fromstring(styles_xml)
        document_root = ET.fromstring(document_xml)
    except ET.ParseError as exc:
        raise DocxCleanerError(f"failed to parse Word XML: {exc}") from exc

    processor = _AllModeProcessor(styles_root)
    changed_paragraphs, direct_outline_removed = processor.clean_document(document_root)

    new_styles_xml = _serialize_xml(styles_root)
    new_document_xml = _serialize_xml(document_root)

    report = CleanReport(
        input_path=source,
        output_path=target,
        changed_paragraphs=changed_paragraphs,
        cloned_styles=processor.cloned_styles,
        direct_outline_removed=direct_outline_removed,
    )

    _write_docx_copy(
        source=source,
        target=target,
        replacements={
            "word/styles.xml": new_styles_xml,
            "word/document.xml": new_document_xml,
        },
    )
    return report


class _AllModeProcessor:
    def __init__(self, styles_root: ET.Element) -> None:
        self.styles_root = styles_root
        self.styles_by_id = self._build_style_map()
        self.clone_by_original: dict[str, str] = {}
        self.nav_cache: dict[str, bool] = {}
        self.cloned_styles = 0

    def clean_document(self, document_root: ET.Element) -> tuple[int, int]:
        changed_paragraphs = 0
        direct_outline_removed = 0

        for paragraph in document_root.iter(f"{W}p"):
            changed = False
            p_pr = paragraph.find(f"{W}pPr")
            if p_pr is None:
                continue

            p_style = p_pr.find(f"{W}pStyle")
            if p_style is not None:
                style_id = p_style.get(f"{W_ATTR}val")
                if style_id and self._style_is_navigation_style(style_id):
                    p_style.set(f"{W_ATTR}val", self._clone_no_nav_style(style_id))
                    changed = True

            outline = p_pr.find(f"{W}outlineLvl")
            if _is_navigation_outline(outline):
                p_pr.remove(outline)
                direct_outline_removed += 1
                changed = True

            if changed:
                changed_paragraphs += 1

        return changed_paragraphs, direct_outline_removed

    def _build_style_map(self) -> dict[str, ET.Element]:
        styles: dict[str, ET.Element] = {}
        for style in self.styles_root.findall(f"{W}style"):
            style_id = style.get(f"{W_ATTR}styleId")
            if style_id:
                styles[style_id] = style
        return styles

    def _style_is_navigation_style(self, style_id: str, seen: set[str] | None = None) -> bool:
        if style_id in self.nav_cache:
            return self.nav_cache[style_id]

        if seen is None:
            seen = set()
        if style_id in seen:
            self.nav_cache[style_id] = False
            return False
        seen.add(style_id)

        style = self.styles_by_id.get(style_id)
        if style is None or style.get(f"{W_ATTR}type") not in (None, "paragraph"):
            self.nav_cache[style_id] = False
            return False

        p_pr = style.find(f"{W}pPr")
        if p_pr is not None and _is_navigation_outline(p_pr.find(f"{W}outlineLvl")):
            self.nav_cache[style_id] = True
            return True

        base_id = _based_on_style_id(style)
        result = bool(base_id and self._style_is_navigation_style(base_id, seen))
        self.nav_cache[style_id] = result
        return result

    def _clone_no_nav_style(self, style_id: str) -> str:
        if style_id in self.clone_by_original:
            return self.clone_by_original[style_id]

        existing_clone_id = self._existing_usable_clone_id(style_id)
        if existing_clone_id is not None:
            self.clone_by_original[style_id] = existing_clone_id
            return existing_clone_id

        original = self.styles_by_id.get(style_id)
        if original is None:
            raise DocxCleanerError(f"style not found while cloning: {style_id}")

        new_style_id = self._new_style_id(style_id)
        clone = copy.deepcopy(original)
        clone.set(f"{W_ATTR}styleId", new_style_id)
        clone.set(f"{W_ATTR}customStyle", "1")

        name = clone.find(f"{W}name")
        if name is None:
            name = ET.Element(f"{W}name")
            clone.insert(0, name)
        name.set(f"{W_ATTR}val", _no_nav_style_name(_style_name(original), style_id))

        base = clone.find(f"{W}basedOn")
        base_id = base.get(f"{W_ATTR}val") if base is not None else None
        if base_id == style_id:
            clone.remove(base)
        elif base is not None and base_id and self._style_is_navigation_style(base_id):
            base.set(f"{W_ATTR}val", self._clone_no_nav_style(base_id))

        p_pr = clone.find(f"{W}pPr")
        if p_pr is not None:
            outline = p_pr.find(f"{W}outlineLvl")
            if outline is not None:
                p_pr.remove(outline)

        self.styles_root.append(clone)
        self.styles_by_id[new_style_id] = clone
        self.clone_by_original[style_id] = new_style_id
        self.nav_cache.clear()
        self.cloned_styles += 1
        return new_style_id

    def _existing_usable_clone_id(self, style_id: str) -> str | None:
        preferred_ids = [f"{style_id}NoNav", f"{style_id}_NoNav"]
        for clone_id in preferred_ids:
            if clone_id in self.styles_by_id and not self._style_is_navigation_style(clone_id):
                return clone_id
        return None

    def _new_style_id(self, style_id: str) -> str:
        base = _safe_style_id(f"{style_id}NoNav")
        candidate = base
        index = 2
        while candidate in self.styles_by_id:
            candidate = f"{base}{index}"
            index += 1
        return candidate


def _validate_paths(source: Path, target: Path, overwrite: bool = False) -> None:
    if not source.exists():
        raise DocxCleanerError(f"input file does not exist: {source}")
    if not source.is_file():
        raise DocxCleanerError(f"input is not a file: {source}")
    if source.suffix.lower() != ".docx":
        raise DocxCleanerError("only .docx files are supported")
    if target.suffix.lower() != ".docx":
        raise DocxCleanerError("output path must end with .docx")
    if source == target:
        raise DocxCleanerError("refusing to overwrite the input file")
    if target.exists() and not overwrite:
        raise DocxCleanerError(f"output file already exists: {target}")
    if not target.parent.exists():
        raise DocxCleanerError(f"output directory does not exist: {target.parent}")


def _require_docx_parts(names: set[str]) -> None:
    missing = [name for name in ("word/styles.xml", "word/document.xml") if name not in names]
    if missing:
        joined = ", ".join(missing)
        raise DocxCleanerError(f"missing required .docx part(s): {joined}")


def _is_navigation_outline(outline: ET.Element | None) -> bool:
    if outline is None:
        return False
    value = outline.get(f"{W_ATTR}val")
    if value is None:
        return True
    return value.isdigit() and 0 <= int(value) <= 8


def _based_on_style_id(style: ET.Element) -> str | None:
    based_on = style.find(f"{W}basedOn")
    if based_on is None:
        return None
    return based_on.get(f"{W_ATTR}val")


def _style_name(style: ET.Element) -> str | None:
    name = style.find(f"{W}name")
    if name is None:
        return None
    return name.get(f"{W_ATTR}val")


def _no_nav_style_name(name: str | None, fallback_style_id: str) -> str:
    base = name or fallback_style_id
    suffix = "（不进导航）" if _contains_cjk(base) else " (No Nav)"
    if base.endswith(suffix):
        return base
    return f"{base}{suffix}"


def _contains_cjk(text: str) -> bool:
    return any("\u4e00" <= char <= "\u9fff" for char in text)


def _safe_style_id(style_id: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "", style_id)
    return cleaned or "NoNavStyle"


def _serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True, short_empty_elements=True)


def _write_docx_copy(source: Path, target: Path, replacements: dict[str, bytes]) -> None:
    temp_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{target.name}.", suffix=".tmp", dir=target.parent, delete=False
        ) as tmp:
            temp_name = tmp.name

        with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(temp_name, "w") as zout:
            for item in zin.infolist():
                data = replacements.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                _copy_zip_entry(zout, item, data)

        os.replace(temp_name, target)
        temp_name = None
    finally:
        if temp_name and os.path.exists(temp_name):
            os.unlink(temp_name)


def _copy_zip_entry(zout: zipfile.ZipFile, item: zipfile.ZipInfo, data: bytes) -> None:
    new_info = zipfile.ZipInfo(filename=item.filename, date_time=item.date_time)
    new_info.comment = item.comment
    new_info.extra = item.extra
    new_info.internal_attr = item.internal_attr
    new_info.external_attr = item.external_attr
    new_info.create_system = item.create_system
    new_info.compress_type = item.compress_type
    new_info._compresslevel = getattr(item, "_compresslevel", None)
    zout.writestr(new_info, data)


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx-heading-cleaner",
        description="Clean Word navigation headings while preserving visible formatting.",
    )
    parser.add_argument("input", help="input .docx file")
    parser.add_argument("--mode", default="all", choices=["all"], help="cleaning mode; only all is supported")
    parser.add_argument("--out", help="output .docx file; defaults to input.cleaned.docx")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    parser = _build_arg_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)

    try:
        report = clean_docx_all(args.input, args.out)
    except DocxCleanerError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    print(report.to_text())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
