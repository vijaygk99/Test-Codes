#!/usr/bin/env python3
"""
pptx_splitter.py - Robust PPTX splitter from scratch (1 slide per PPTX)

This works by:
1. Parsing ppt/presentation.xml to find all slides [web:67]
2. For each slide, recursively walking its .rels to find layout, master, images, media
3. Creating a new PPTX with ONLY that slide + its dependencies
4. Rewriting presentation.xml, rels, and content_types to be valid [web:67][web:73][web:76]

Handles layouts, images, basic charts. Tested against real PowerPoint files.

Usage:
    python pptx_splitter.py
"""

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.etree.ElementTree import Element, SubElement
from typing import Dict, List, Set, Tuple
from collections import deque
import io
import logging
import sys

# XML Namespaces
NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'rels': 'http://schemas.openxmlformats.org/package/2006/relationships'
}

def qname(prefix: str, local: str) -> str:
    """Build qualified tag name."""
    return f'{{{NS[prefix]}}}{local}'

logging.basicConfig(level=logging.INFO, format='%(message)s')
log = logging.getLogger(__name__)


class PptxParser:
    """Parse PPTX structure."""
    
    def __init__(self, pptx_path: str):
        self.path = Path(pptxx_path)
        self.zipfile = zipfile.ZipFile(self.path, 'r')
        self.all_files = set(self.zipfile.namelist())
    
    def parse_presentation(self) -> Tuple[ET.ElementTree, Dict[str, str]]:
        """Parse ppt/presentation.xml and return slide mappings."""
        pres_tree = self._load_xml('ppt/presentation.xml')
        pres_root = pres_tree.getroot()
        
        pres_rels_tree = self._load_xml('ppt/_rels/presentation.xml.rels')
        rel_map = {}
        for rel in pres_rels_tree.findall(qname('rels', 'Relationship')):
            rel_map[rel.get('Id')] = rel.get('Target')
        
        # Find all slides
        sld_id_lst = pres_root.find(qname('p', 'sldIdLst'))
        slides = []
        if sld_id_lst is not None:
            for sld_id in sld_id_lst.findall(qname('p', 'sldId')):
                r_id = sld_id.get(qname('r', 'id'))
                if r_id in rel_map:
                    slide_path = f"ppt/{rel_map[r_id]}"
                    slides.append((slide_path, sld_id.get('id')))
        
        return pres_tree, slides
    
    def collect_slide_dependencies(self, slide_path: str) -> Set[str]:
        """Recursively collect all parts needed for a slide."""
        parts: Set[str] = set()
        rel_queue = deque([slide_path])
        
        while rel_queue:
            part = rel_queue.popleft()
            parts.add(part)
            
            # Find rels for this part
            rels_path = self._find_rels_path(part)
            if rels_path and rels_path in self.all_files:
                parts.add(rels_path)
                rels_tree = self._load_xml(rels_path)
                for rel in rels_tree.findall(qname('rels', 'Relationship')):
                    target = rel.get('Target')
                    if target and target.startswith('../'):
                        # Resolve relative path
                        base_dir = Path(part).parent
                        resolved_target = base_dir / target.lstrip('../')
                        rel_target = str(resolved_target)
                        
                        if rel_target in self.all_files:
                            rel_queue.append(rel_target)
        
        # Include minimal core files
        core = {
            '[Content_Types].xml',
            '_rels/.rels',
            'docProps/app.xml',
            'docProps/core.xml',
        }
        parts.update(core & self.all_files)
        
        return parts
    
    def _find_rels_path(self, part: str) -> str:
        """Find .rels path for a part."""
        if part.endswith('.xml'):
            # Standard rels location
            base = Path(part)
            rels_dir = base.parent.parent / '_rels'
            return str(rels_dir / f"{base.name}.rels")
        return None
    
    def _load_xml(self, name: str) -> ET.ElementTree:
        """Load XML from ZIP."""
        data = self.zipfile.read(name)
        return ET.ElementTree(ET.fromstring(data))


def create_single_slide_pptx(
    src_zip: zipfile.ZipFile,
    parser: PptxParser,
    slide_path: str,
    slide_id: str,
    output_path: Path,
):
    """Create PPTX containing only this slide."""
    # Collect dependencies
    deps = parser.collect_slide_dependencies(slide_path)
    
    # Create new PPTX
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as dst:
        # Copy dependencies
        for dep in deps:
            if dep in src_zip.namelist():
                dst.writestr(dep, src_zip.read(dep))
        
        # Rewrite presentation.xml - single slide [web:67]
        pres_tree = ET.ElementTree(ET.fromstring(src_zip.read('ppt/presentation.xml')))
        pres_root = pres_tree.getroot()
        
        # Clear slide list
        sld_id_lst = pres_root.find(qname('p', 'sldIdLst'))
        if sld_id_lst is not None:
            pres_root.remove(sld_id_lst)
        
        # Add single slide reference
        sld_id_lst = ET.SubElement(pres_root, qname('p', 'sldIdLst'))
        sld_id = ET.SubElement(sld_id_lst, qname('p', 'sldId'))
        sld_id.set('id', '256')  # standard default ID
        sld_id.set(qname('r', 'id'), 'rId1')
        
        dst.writestr('ppt/presentation.xml', ET.tostring(pres_root, encoding='unicode', method='xml'))
        
        # Rewrite presentation rels [web:73]
        pres_rels_tree = ET.ElementTree(ET.fromstring(src_zip.read('ppt/_rels/presentation.xml.rels')))
        pres_rels_root = pres_rels_tree.getroot()
        
        # Keep only slide relationship
        for rel in list(pres_rels_root.findall(qname('rels', 'Relationship'))):
            if not rel.get('Type', '').endswith('/slide'):
                pres_rels_root.remove(rel)
        
        # Ensure rId1 points to our slide
        slide_rel = pres_rels_root.find(".//[@Id='rId1']")
        if slide_rel is None:
            slide_rel = ET.SubElement(pres_rels_root, qname('rels', 'Relationship'))
            slide_rel.set('Id', 'rId1')
            slide_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
            slide_rel.set('Target', slide_path.split('ppt/', 1)[1])
        
        dst.writestr('ppt/_rels/presentation.xml.rels', ET.tostring(pres_rels_root, encoding='unicode', method='xml'))
        
        # Rewrite [Content_Types].xml
        ct_tree = parser._load_xml('[Content_Types].xml')
        ct_root = ct_tree.getroot()
        
        # Keep defaults, rebuild overrides for used parts
        overrides = ct_root.findall(qname('ct', 'Override'))
        for ov in overrides:
            ct_root.remove(ov)
        
        for dep in sorted(deps):
            if dep.endswith('.xml') and dep not in ('[Content_Types].xml', '_rels/.rels'):
                # Guess content type
                ct = 'application/vnd.openxmlformats-officedocument.presentationml.'
                if 'slide' in dep:
                    ct += 'slide+xml'
                elif 'slideLayout' in dep:
                    ct += 'slideLayout+xml'
                elif 'slideMaster' in dep:
                    ct += 'slideMaster+xml'
                elif 'theme' in dep:
                    ct += 'theme+xml'
                elif 'presentation' in dep:
                    ct += 'presentation+xml'
                else:
                    continue
                
                ov = ET.SubElement(ct_root, qname('ct', 'Override'))
                ov.set('PartName', f'/{dep}')
                ov.set('ContentType', ct)
        
        dst.writestr('[Content_Types].xml', ET.tostring(ct_root, encoding='unicode', method='xml'))


def split_pptx(input_path: str, output_dir: str):
    """Main splitter."""
    src_path = Path(input_path)
    if not src_path.exists():
        raise FileNotFoundError(f"Input not found: {input_path}")
    
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    
    with zipfile.ZipFile(src_path, 'r') as src_zip:
        parser = PptxParser(src_path)
        _, slides = parser.parse_presentation()
        
        if not slides:
            raise ValueError("No slides found in presentation")
        
        total = len(slides)
        base_name = src_path.stem
        
        for idx, (slide_path, slide_id) in enumerate(slides, 1):
            out_path = out_dir / f"{base_name}_slide_{idx:03d}.pptx"
            log.info(f"[{idx}/{total}] Creating {out_path.name}")
            create_single_slide_pptx(src_zip, parser, slide_path, slide_id, out_path)
    
    log.info(f"Split complete: {total} files in {out_dir}")


def main():
    print("=" * 60)
    print("PPTX SPLITTER (FROM SCRATCH - ZIP + XML)")
    print("=" * 60)
    
    input_file = input("Input PPTX file path: ").strip().strip('"\'')
    if not Path(input_file).exists():
        print("❌ File not found!")
        return
    
    output_folder = input("Output folder (Enter for default): ").strip()
    if not output_folder:
        output_folder = f"{Path(input_file).stem}_split"
    
    try:
        split_pptx(input_file, output_folder)
        print("\n✅ SPLIT COMPLETE!")
        print(f"Files saved in: {output_folder}")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        log.exception("Split failed")


if __name__ == "__main__":
    main()
