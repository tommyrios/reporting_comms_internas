from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

REL_TYPE_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
REL_TYPE_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"

ET.register_namespace("", CT_NS)
ET.register_namespace("", REL_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)


@dataclass
class PptxPackage:
    files: dict[str, bytes]

    @classmethod
    def from_path(cls, path: Path) -> "PptxPackage":
        with zipfile.ZipFile(path, "r") as zf:
            return cls({name: zf.read(name) for name in zf.namelist()})

    def to_path(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for name in sorted(self.files.keys()):
                zf.writestr(name, self.files[name])


def _rels_path(part_path: str) -> str:
    directory, filename = posixpath.split(part_path)
    if directory:
        return f"{directory}/_rels/{filename}.rels"
    return f"_rels/{filename}.rels"


def _resolve_target(source_part: str, target: str) -> str:
    base_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base_dir, target))


def _relative_target(from_part: str, to_part: str) -> str:
    from_dir = posixpath.dirname(from_part)
    return posixpath.relpath(to_part, start=from_dir or ".")


def _xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def _xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _next_rid(rels_root: ET.Element) -> str:
    max_id = 0
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.attrib.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            max_id = max(max_id, int(rid[3:]))
    return f"rId{max_id + 1}"


def _next_slide_id(sld_id_lst: ET.Element) -> str:
    max_id = 255
    for node in sld_id_lst.findall(f"{{{P_NS}}}sldId"):
        raw = node.attrib.get("id", "")
        if raw.isdigit():
            max_id = max(max_id, int(raw))
    return str(max_id + 1)


def _next_numeric_part(existing: set[str], desired: str) -> str:
    if desired not in existing:
        return desired

    directory, filename = posixpath.split(desired)
    match = re.match(r"^(.*?)(\d+)(\.[^.]+)$", filename)
    if match:
        prefix, _, suffix = match.groups()
        used = []
        regex = re.compile(rf"^{re.escape(prefix)}(\d+){re.escape(suffix)}$")
        for part in existing:
            d, f = posixpath.split(part)
            if d != directory:
                continue
            m = regex.match(f)
            if m:
                used.append(int(m.group(1)))
        next_n = (max(used) + 1) if used else 1
        return f"{directory}/{prefix}{next_n}{suffix}" if directory else f"{prefix}{next_n}{suffix}"

    stem, ext = posixpath.splitext(filename)
    counter = 1
    while True:
        candidate = f"{directory}/{stem}_copy{counter}{ext}" if directory else f"{stem}_copy{counter}{ext}"
        if candidate not in existing:
            return candidate
        counter += 1


@dataclass
class PartCopier:
    source: PptxPackage
    target: PptxPackage
    content_types_src: ET.Element
    content_types_dst: ET.Element
    part_map: dict[str, str] = field(default_factory=dict)

    def _ensure_content_type(self, source_part: str, target_part: str) -> None:
        source_override = self.content_types_src.find(f"{{{CT_NS}}}Override[@PartName='/{source_part}']")
        if source_override is not None:
            target_override = self.content_types_dst.find(f"{{{CT_NS}}}Override[@PartName='/{target_part}']")
            if target_override is None:
                ET.SubElement(
                    self.content_types_dst,
                    f"{{{CT_NS}}}Override",
                    {
                        "PartName": f"/{target_part}",
                        "ContentType": source_override.attrib.get("ContentType", "application/octet-stream"),
                    },
                )

        ext = posixpath.splitext(target_part)[1].lstrip(".").lower()
        if ext:
            source_default = self.content_types_src.find(f"{{{CT_NS}}}Default[@Extension='{ext}']")
            target_default = self.content_types_dst.find(f"{{{CT_NS}}}Default[@Extension='{ext}']")
            if source_default is not None and target_default is None:
                ET.SubElement(
                    self.content_types_dst,
                    f"{{{CT_NS}}}Default",
                    {
                        "Extension": ext,
                        "ContentType": source_default.attrib.get("ContentType", "application/octet-stream"),
                    },
                )

    def copy_part(self, source_part: str, xml_replacements: dict[str, str] | None = None) -> str:
        if source_part in self.part_map:
            return self.part_map[source_part]

        existing = set(self.target.files.keys())
        target_part = _next_numeric_part(existing, source_part)
        self.part_map[source_part] = target_part

        payload = self.source.files[source_part]
        if xml_replacements and source_part.endswith(".xml"):
            try:
                text = payload.decode("utf-8")
                for old, new in xml_replacements.items():
                    text = text.replace(old, new)
                payload = text.encode("utf-8")
            except Exception:
                pass
        self.target.files[target_part] = payload
        self._ensure_content_type(source_part, target_part)

        source_rels_path = _rels_path(source_part)
        if source_rels_path in self.source.files:
            rels_root = _xml(self.source.files[source_rels_path])
            for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
                if rel.attrib.get("TargetMode") == "External":
                    continue
                target = rel.attrib.get("Target", "")
                if not target:
                    continue
                nested_source_part = _resolve_target(source_part, target)
                if nested_source_part not in self.source.files:
                    continue
                nested_target_part = self.copy_part(nested_source_part)
                rel.attrib["Target"] = _relative_target(target_part, nested_target_part)

            self.target.files[_rels_path(target_part)] = _xml_bytes(rels_root)

        return target_part


def _slide_parts_from_presentation(pkg: PptxPackage) -> list[str]:
    presentation = _xml(pkg.files["ppt/presentation.xml"])
    rels = _xml(pkg.files["ppt/_rels/presentation.xml.rels"])

    rel_map = {
        rel.attrib.get("Id"): _resolve_target("ppt/presentation.xml", rel.attrib.get("Target", ""))
        for rel in rels.findall(f"{{{REL_NS}}}Relationship")
        if rel.attrib.get("Type") == REL_TYPE_SLIDE
    }

    slides: list[str] = []
    sld_id_lst = presentation.find(f"{{{P_NS}}}sldIdLst")
    if sld_id_lst is None:
        return slides
    for node in sld_id_lst.findall(f"{{{P_NS}}}sldId"):
        rid = node.attrib.get(f"{{{R_NS}}}id")
        if rid in rel_map:
            slides.append(rel_map[rid])
    return slides


def _append_master_refs_if_missing(pkg: PptxPackage, copied_parts: list[str]) -> None:
    master_parts = sorted({part for part in copied_parts if part.startswith("ppt/slideMasters/") and part.endswith(".xml")})
    if not master_parts:
        return

    presentation = _xml(pkg.files["ppt/presentation.xml"])
    rels = _xml(pkg.files["ppt/_rels/presentation.xml.rels"])

    target_to_rid = {
        _resolve_target("ppt/presentation.xml", rel.attrib.get("Target", "")): rel.attrib.get("Id")
        for rel in rels.findall(f"{{{REL_NS}}}Relationship")
    }

    sld_master_id_lst = presentation.find(f"{{{P_NS}}}sldMasterIdLst")
    if sld_master_id_lst is None:
        sld_master_id_lst = ET.Element(f"{{{P_NS}}}sldMasterIdLst")
        presentation.insert(0, sld_master_id_lst)

    existing_master_rids = {
        node.attrib.get(f"{{{R_NS}}}id")
        for node in sld_master_id_lst.findall(f"{{{P_NS}}}sldMasterId")
    }

    next_master_id = 2**31 - 10
    for node in sld_master_id_lst.findall(f"{{{P_NS}}}sldMasterId"):
        raw = node.attrib.get("id", "")
        if raw.isdigit():
            next_master_id = max(next_master_id, int(raw))

    for master_part in master_parts:
        rid = target_to_rid.get(master_part)
        if not rid:
            rid = _next_rid(rels)
            ET.SubElement(
                rels,
                f"{{{REL_NS}}}Relationship",
                {
                    "Id": rid,
                    "Type": REL_TYPE_SLIDE_MASTER,
                    "Target": _relative_target("ppt/presentation.xml", master_part),
                },
            )
            target_to_rid[master_part] = rid

        if rid not in existing_master_rids:
            next_master_id += 1
            ET.SubElement(
                sld_master_id_lst,
                f"{{{P_NS}}}sldMasterId",
                {"id": str(next_master_id), f"{{{R_NS}}}id": rid},
            )
            existing_master_rids.add(rid)

    pkg.files["ppt/presentation.xml"] = _xml_bytes(presentation)
    pkg.files["ppt/_rels/presentation.xml.rels"] = _xml_bytes(rels)


def _insert_slide_reference(pkg: PptxPackage, slide_part: str, prepend: bool) -> None:
    presentation = _xml(pkg.files["ppt/presentation.xml"])
    rels = _xml(pkg.files["ppt/_rels/presentation.xml.rels"])

    slide_rel = ET.SubElement(
        rels,
        f"{{{REL_NS}}}Relationship",
        {
            "Id": _next_rid(rels),
            "Type": REL_TYPE_SLIDE,
            "Target": _relative_target("ppt/presentation.xml", slide_part),
        },
    )

    sld_id_lst = presentation.find(f"{{{P_NS}}}sldIdLst")
    if sld_id_lst is None:
        raise RuntimeError("presentation.xml sin sldIdLst")

    node = ET.Element(
        f"{{{P_NS}}}sldId",
        {
            "id": _next_slide_id(sld_id_lst),
            f"{{{R_NS}}}id": slide_rel.attrib["Id"],
        },
    )

    if prepend:
        sld_id_lst.insert(0, node)
    else:
        sld_id_lst.append(node)

    pkg.files["ppt/presentation.xml"] = _xml_bytes(presentation)
    pkg.files["ppt/_rels/presentation.xml.rels"] = _xml_bytes(rels)


def assemble_deck(template_path: Path, body_path: Path, output_path: Path, period_label: str) -> None:
    template_pkg = PptxPackage.from_path(Path(template_path))
    body_pkg = PptxPackage.from_path(Path(body_path))

    template_slide_parts = _slide_parts_from_presentation(template_pkg)
    if len(template_slide_parts) < 2:
        raise RuntimeError("La plantilla debe tener al menos 2 slides (portada y cierre)")

    cover_part = template_slide_parts[0]
    closing_part = template_slide_parts[-1]

    content_types_src = _xml(template_pkg.files["[Content_Types].xml"])
    content_types_dst = _xml(body_pkg.files["[Content_Types].xml"])

    copier = PartCopier(
        source=template_pkg,
        target=body_pkg,
        content_types_src=content_types_src,
        content_types_dst=content_types_dst,
    )

    copied_cover = copier.copy_part(cover_part, xml_replacements={"FECHA": str(period_label or "-")})
    copied_closing = copier.copy_part(closing_part)

    _append_master_refs_if_missing(body_pkg, list(copier.part_map.values()))
    _insert_slide_reference(body_pkg, copied_cover, prepend=True)
    _insert_slide_reference(body_pkg, copied_closing, prepend=False)

    body_pkg.files["[Content_Types].xml"] = _xml_bytes(content_types_dst)
    body_pkg.to_path(Path(output_path))


__all__ = ["assemble_deck"]
