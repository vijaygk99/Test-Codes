"""
pptx_split_from_scratch.py

Split a PPTX into single-slide PPTX files using low-level ZIP + XML.
This avoids python-pptx cloning and works directly on the Open XML package. [web:71][web:73]

STATUS: Prototype focused on typical decks (text, images, standard layouts).
Complex charts/SmartArt/embedded OLE may need more part types added
to the dependency collector.

Usage:
    python pptx_split_from_scratch.py

It will prompt:
    - Input PPTX path
    - Output folder
"""

import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET  # [web:84]
import shutil
import io


# Namespaces used in PPTX XML
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Convenience tag builder
def qname(prefix, local):
    return f"{{{NS[prefix]}}}{local}"


def load_xml_from_zip(zf: zipfile.ZipFile, path: str) -> ET.ElementTree:
    data = zf.read(path)
    return ET.ElementTree(ET.fromstring(data))


def write_xml_to_zip(zf: zipfile.ZipFile, path: str, tree: ET.ElementTree):
    buf = io.BytesIO()
    tree.write(buf, encoding="utf-8", xml_declaration=True)
    zf.writestr(path, buf.getvalue())


def get_presentation_info(zf: zipfile.ZipFile):
    """Return (presentation_tree, pres_rels_tree, slide_id_elems, rels_map)."""
    # main presentation part
    pres_path = "ppt/presentation.xml"
    pres_tree = load_xml_from_zip(zf, pres_path)
    pres_root = pres_tree.getroot()

    # its relationships part [web:73][web:76]
    pres_rels_path = "ppt/_rels/presentation.xml.rels"
    pres_rels_tree = load_xml_from_zip(zf, pres_rels_path)
    pres_rels_root = pres_rels_tree.getroot()

    # Map rId -> slide target path
    rels_map = {}
    for rel in pres_rels_root.findall("Relationship", {"": "http://schemas.openxmlformats.org/package/2006/relationships"}):
        rId = rel.get("Id")
        target = rel.get("Target")
        rels_map[rId] = target  # e.g., "slides/slide1.xml"

    # List slideId elements in order [web:67]
    slide_id_lst = pres_root.find(qname("p", "sldIdLst"))
    slide_ids = [] if slide_id_lst is None else list(slide_id_lst.findall(qname("p", "sldId")))
    return pres_tree, pres_rels_tree, slide_ids, rels_map


def collect_slide_dependencies(zf: zipfile.ZipFile, slide_target: str):
    """
    Given a slide target like 'slides/slide1.xml', collect all part paths
    that are needed for that slide:
      - the slide itself
      - its .rels
      - referenced images/media/charts/notes
      - the slide layout and master & their themes (basic). [web:67][web:73]
    """
    parts = set()
    rels_to_visit = []

    # normalize slide path
    slide_path = f"ppt/{slide_target}"
    parts.add(slide_path)

    # slide rels
    slide_rels_path = f"ppt/slides/_rels/{Path(slide_target).name}.rels"
    if slide_rels_path in zf.namelist():
        parts.add(slide_rels_path)
        rels_to_visit.append(slide_rels_path)

    # BFS over relationships for internal targets
    while rels_to_visit:
        rels_path = rels_to_visit.pop()
        rels_tree = load_xml_from_zip(zf, rels_path)
        rels_root = rels_tree.getroot()
        for rel in rels_root.findall("Relationship", {"": "http://schemas.openxmlformats.org/package/2006/relationships"}):
            target = rel.get("Target")
            mode = rel.get("TargetMode")
            if mode == "External":
                continue  # skip external links
            # build absolute-ish path relative to rels_path dir
            base_dir = Path(rels_path).parent  # e.g., ppt/slides/_rels
            # Most slide rel targets are like "../media/image1.png" or "../slideLayouts/slideLayout1.xml"
            target_path = (base_dir / target).resolve().as_posix()
            # normalize to strip leading "../../"
            # quick normalization: collapse 'ppt/../' patterns
            parts.add(target_path)

            # queue nested rels (e.g., chart.xml.rels) [web:73]
            rels_candidate = f"{target_path}.rels"
            if rels_candidate in zf.namelist() and rels_candidate not in parts:
                parts.add(rels_candidate)
                rels_to_visit.append(rels_candidate)

    # minimal presentation-level dependencies:
    # - we will reuse original slideMaster/slideLayout/theme parts from root,
    #   but for a prototype, include all masters/layouts to be safe.
    for name in zf.namelist():
        if name.startswith("ppt/slideMasters/") or name.startswith("ppt/slideLayouts/") or name.startswith("ppt/theme/"):
            parts.add(name)

    return parts


def build_minimal_content_types(zf: zipfile.ZipFile, used_parts: set[str]):
    """
    Build a [Content_Types].xml that includes only entries needed for used_parts. [web:76][web:79]
    """
    # Load original
    ct_tree = load_xml_from_zip(zf, "[Content_Types].xml")
    ct_root = ct_tree.getroot()

    # Remove all existing Override elements first
    for override in list(ct_root.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override")):
        ct_root.remove(override)

    # Keep Defaults (extensions) as-is (they're generic)
    # Add Override only for parts we actually include
    for part_name in sorted(used_parts):
        # Find matching override in original
        orig_tree = load_xml_from_zip(zf, "[Content_Types].xml")
        orig_root = orig_tree.getroot()
        found = None
        for override in orig_root.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override"):
            if override.get("PartName") == f"/{part_name}":
                found = override
                break
        if found is None:
            # fallback: guess common content types
            # basic heuristics:
            if part_name.endswith(".xml"):
                ctype = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
            elif part_name.endswith(".png"):
                ctype = "image/png"
            elif part_name.endswith(".jpg") or part_name.endswith(".jpeg"):
                ctype = "image/jpeg"
            else:
                # skip unknown for now
                continue
        else:
            ctype = found.get("ContentType")

        new_override = ET.Element(
            "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
            PartName=f"/{part_name}",
            ContentType=ctype,
        )
        ct_root.append(new_override)

    return ct_tree


def create_single_slide_pptx(
    src_zip: zipfile.ZipFile,
    slide_idx: int,
    slide_id_elem,
    rels_map: dict,
    output_path: Path,
):
    """
    Create a new PPTX containing only the given slide.
    """
    # 1) identify slide target path from slideId's r:id [web:67]
    rId = slide_id_elem.get(f"{{{NS['r']}}}id")
    slide_target = rels_map[rId]  # e.g., "slides/slide1.xml"

    # 2) collect all dependent parts for this slide
    used_parts = collect_slide_dependencies(src_zip, slide_target)

    # 3) always include core parts
    core_parts = {
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/core.xml",
        "docProps/app.xml",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
    }
    used_parts.update(p for p in core_parts if p in src_zip.namelist())

    # 4) construct new ZIP
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as dst_zf:
        # Copy used parts as-is for now
        for name in used_parts:
            dst_zf.writestr(name, src_zip.read(name))

        # Rewrite ppt/presentation.xml to reference only this slide [web:67]
        pres_tree = load_xml_from_zip(dst_zf, "ppt/presentation.xml")
        pres_root = pres_tree.getroot()
        sldIdLst = pres_root.find(qname("p", "sldIdLst"))
        if sldIdLst is None:
            sldIdLst = ET.SubElement(pres_root, qname("p", "sldIdLst"))
        # Remove all existing sldId, then add only this one
        for old in list(sldIdLst.findall(qname("p", "sldId"))):
            sldIdLst.remove(old)

        # Clone the slideId element but reset id to something simple
        new_sldId = ET.Element(qname("p", "sldId"))
        new_sldId.set("id", "256")  # arbitrary valid ID [web:67]
        new_sldId.set(f"{{{NS['r']}}}id", "rId1")
        sldIdLst.append(new_sldId)

        write_xml_to_zip(dst_zf, "ppt/presentation.xml", pres_tree)

        # Rewrite ppt/_rels/presentation.xml.rels to point rId1 to our slide [web:73]
        pres_rels_tree = load_xml_from_zip(dst_zf, "ppt/_rels/presentation.xml.rels")
        pres_rels_root = pres_rels_tree.getroot()
        # Remove all existing slide relationships
        for rel in list(pres_rels_root.findall("Relationship", {"": "http://schemas.openxmlformats.org/package/2006/relationships"})):
            if rel.get("Type", "").endswith("/slide"):
                pres_rels_root.remove(rel)
        # Add new rel rId1 -> slide target
        ET.SubElement(
            pres_rels_root,
            "Relationship",
            {
                "Id": "rId1",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                "Target": slide_target,
            },
        )
        write_xml_to_zip(dst_zf, "ppt/_rels/presentation.xml.rels", pres_rels_tree)

        # Rewrite [Content_Types].xml for only used_parts [web:79]
        ct_tree = build_minimal_content_types(src_zip, used_parts - {"[Content_Types].xml"})
        write_xml_to_zip(dst_zf, "[Content_Types].xml", ct_tree)


def split_pptx_file(input_pptx: str, output_dir: str):
    src_path = Path(input_pptx)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src_path, "r") as zf:
        pres_tree, pres_rels_tree, slide_ids, rels_map = get_presentation_info(zf)
        total = len(slide_ids)
        base = src_path.stem
        for idx, sldId in enumerate(slide_ids, start=1):
            out_name = out_dir / f"{base}_slide_{idx:03d}.pptx"
            print(f"Creating {out_name.name} ({idx}/{total}) ...")
            create_single_slide_pptx(zf, idx, sldId, rels_map, out_name)

    print(f"\nDone. Created {total} files in {out_dir}")


def main():
    print("=" * 60)
    print(" PPTX SPLITTER (LOW-LEVEL OOXML PROTOTYPE)")
    print("=" * 60)
    in_path = input("Enter input PPTX path: ").strip().strip('"').strip("'")
    if not in_path:
        print("No input provided, exiting.")
        return
    out_default = str(Path(in_path).with_suffix("")) + "_split"
    out_path = input(f"Enter output folder (default: {out_default}): ").strip()
    if not out_path:
        out_path = out_default

    split_pptx_file(in_path, out_path)


if __name__ == "__main__":
    main()
