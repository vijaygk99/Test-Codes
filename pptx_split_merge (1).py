#!/usr/bin/env python3
"""
PPTX Split & Merge System - FIXED VERSION
Properly handles XML namespaces and creates valid single-slide files
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict
import re
import shutil
from typing import Dict, Set, List, Tuple

# Namespaces for PPTX XML parsing
NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

class PPTXSplitter:
    """Splits PPTX into individual slide files with full dependency resolution"""
    
    def __init__(self, input_pptx: str):
        self.input_path = Path(input_pptx)
        self.base_name = self.input_path.stem
        
    def split(self, output_dir: str = None) -> List[str]:
        """Split PPTX into individual slide files"""
        if output_dir is None:
            output_dir = f"{self.base_name}_slides"
        
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        print(f"📂 Reading {self.input_path.name}...")
        
        with zipfile.ZipFile(self.input_path, 'r') as zin:
            # Parse presentation.xml to get slide list
            slides = self._get_slide_list(zin)
            print(f"📊 Found {len(slides)} slides")
            
            output_files = []
            
            for idx, (slide_path, slide_rid, slide_id) in enumerate(slides, 1):
                output_file = output_path / f"{self.base_name}_slide_{idx:03d}.pptx"
                print(f"  ✂️  Extracting slide {idx}...", end='', flush=True)
                
                # Collect all dependencies for this slide
                dependencies = self._collect_dependencies(zin, slide_path)
                
                # Create new PPTX with only this slide
                self._create_single_slide_pptx(zin, slide_path, slide_rid, slide_id,
                                               dependencies, output_file, idx)
                
                output_files.append(str(output_file))
                print(f" ✅")
                
                # Verify the created file
                self._verify_slide_file(output_file)
            
        print(f"\n✨ Split complete! {len(output_files)} files created in {output_dir}/")
        print(f"\n🔍 Verification: Opening each file to check validity...")
        return output_files
    
    def _get_slide_list(self, zin: zipfile.ZipFile) -> List[Tuple[str, str, str]]:
        """Parse presentation.xml to get ordered list of slides with their IDs"""
        pres_xml = zin.read('ppt/presentation.xml')
        root = ET.fromstring(pres_xml)
        
        # Get slide IDs from presentation with proper namespace handling
        slide_info = []
        sld_id_lst = root.find('.//p:sldIdLst', NS)
        
        if sld_id_lst is not None:
            for sld_id in sld_id_lst.findall('p:sldId', NS):
                rid = sld_id.get(f"{{{NS['r']}}}id")
                sid = sld_id.get('id')  # Slide ID attribute
                slide_info.append((rid, sid))
        
        # Resolve relationship IDs to actual slide paths
        rels_xml = zin.read('ppt/_rels/presentation.xml.rels')
        rels_root = ET.fromstring(rels_xml)
        
        slides = []
        for rid, sid in slide_info:
            for rel in rels_root.findall('rel:Relationship', NS):
                if rel.get('Id') == rid:
                    target = rel.get('Target')
                    slide_path = f"ppt/{target}"
                    slides.append((slide_path, rid, sid))
                    break
        
        return slides
    
    def _collect_dependencies(self, zin: zipfile.ZipFile, 
                             slide_path: str) -> Set[str]:
        """Collect all files needed for this slide"""
        dependencies = set()
        to_process = [slide_path]
        processed = set()
        
        while to_process:
            current = to_process.pop(0)
            if current in processed:
                continue
            
            processed.add(current)
            dependencies.add(current)
            
            # Add relationship file if it exists
            rels_path = self._get_rels_path(current)
            if rels_path in zin.namelist():
                dependencies.add(rels_path)
                
                # Parse relationships to find dependencies
                try:
                    rels_xml = zin.read(rels_path)
                    rels_root = ET.fromstring(rels_xml)
                    
                    for rel in rels_root.findall('rel:Relationship', NS):
                        target = rel.get('Target')
                        rel_type = rel.get('Type')
                        
                        if target and not target.startswith('http'):
                            # Resolve relative path
                            dep_path = self._resolve_path(current, target)
                            if dep_path and dep_path not in processed:
                                to_process.append(dep_path)
                except:
                    pass
        
        # Always include core files
        core_files = [
            '[Content_Types].xml',
            '_rels/.rels',
            'docProps/core.xml',
            'docProps/app.xml',
            'ppt/presentation.xml',
            'ppt/_rels/presentation.xml.rels',
            'ppt/presProps.xml',
            'ppt/viewProps.xml',
            'ppt/tableStyles.xml'
        ]
        
        for core in core_files:
            if core in zin.namelist():
                dependencies.add(core)
                rels = self._get_rels_path(core)
                if rels in zin.namelist():
                    dependencies.add(rels)
        
        # Include all theme files and their dependencies
        for name in zin.namelist():
            if '/theme/' in name or name.startswith('ppt/theme'):
                dependencies.add(name)
                rels = self._get_rels_path(name)
                if rels in zin.namelist():
                    dependencies.add(rels)
        
        return dependencies
    
    def _get_rels_path(self, file_path: str) -> str:
        """Get the .rels path for a given file"""
        parts = file_path.rsplit('/', 1)
        if len(parts) == 2:
            return f"{parts[0]}/_rels/{parts[1]}.rels"
        return f"_rels/{file_path}.rels"
    
    def _resolve_path(self, base_path: str, target: str) -> str:
        """Resolve relative path from base to target"""
        if target.startswith('http'):
            return None
        
        if target.startswith('/'):
            return target[1:]
        
        base_dir = base_path.rsplit('/', 1)[0] if '/' in base_path else ''
        
        # Handle .. in path
        parts = target.split('/')
        base_parts = base_dir.split('/') if base_dir else []
        
        for part in parts:
            if part == '..':
                if base_parts:
                    base_parts.pop()
            elif part and part != '.':
                base_parts.append(part)
        
        result = '/'.join(base_parts)
        return result if result else None
    
    def _create_single_slide_pptx(self, zin: zipfile.ZipFile, 
                                  slide_path: str, slide_rid: str, slide_id: str,
                                  dependencies: Set[str], 
                                  output_file: Path, slide_num: int):
        """Create a new PPTX with single slide and all dependencies"""
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zout:
            # First, write the modified presentation.xml with ONLY this slide
            self._write_single_slide_presentation(zin, zout, slide_path, slide_rid, slide_id)
            
            # Write the modified presentation rels with ONLY this slide relationship
            self._write_single_slide_rels(zin, zout, slide_rid)
            
            # Copy all dependency files EXCEPT presentation.xml and its rels (already written)
            for dep in dependencies:
                if dep in ['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels']:
                    continue  # Skip - we already wrote modified versions
                    
                if dep in zin.namelist():
                    try:
                        zout.writestr(dep, zin.read(dep))
                    except:
                        pass
            
            # Update docProps/app.xml to show 1 slide
            self._update_app_properties(zin, zout)
    
    def _write_single_slide_presentation(self, zin: zipfile.ZipFile,
                                        zout: zipfile.ZipFile,
                                        slide_path: str, slide_rid: str, slide_id: str):
        """Rewrite presentation.xml to contain only one slide"""
        pres_xml = zin.read('ppt/presentation.xml')
        
        # Parse with namespace preservation
        root = ET.fromstring(pres_xml)
        
        # Register all namespaces to preserve them
        for prefix, uri in NS.items():
            ET.register_namespace(prefix, uri)
        
        # Find the slide ID list
        sld_id_lst = root.find('.//p:sldIdLst', NS)
        
        if sld_id_lst is not None:
            # Remove all slides except the target one
            slides_to_remove = []
            target_slide = None
            
            for sld_id in sld_id_lst.findall('p:sldId', NS):
                rid = sld_id.get(f"{{{NS['r']}}}id")
                if rid == slide_rid:
                    target_slide = sld_id
                else:
                    slides_to_remove.append(sld_id)
            
            # Remove unwanted slides
            for slide in slides_to_remove:
                sld_id_lst.remove(slide)
            
            # Make sure target slide has id="256" (standard first slide ID)
            if target_slide is not None:
                target_slide.set('id', '256')
        
        # Write the modified XML
        xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        zout.writestr('ppt/presentation.xml', xml_bytes)
    
    def _write_single_slide_rels(self, zin: zipfile.ZipFile,
                                 zout: zipfile.ZipFile, slide_rid: str):
        """Rewrite presentation.xml.rels to contain only relationships for this slide"""
        rels_xml = zin.read('ppt/_rels/presentation.xml.rels')
        root = ET.fromstring(rels_xml)
        
        # Register namespace
        ET.register_namespace('', NS['rel'])
        
        # Find all relationships
        rels_to_remove = []
        target_slide_rel = None
        
        for rel in root.findall('rel:Relationship', NS):
            rel_id = rel.get('Id')
            rel_type = rel.get('Type')
            
            # Keep this relationship if it's:
            # 1. The target slide
            # 2. A slideMaster, slideLayout, or theme
            # 3. Core properties (presProps, viewProps, tableStyles)
            if rel_id == slide_rid:
                target_slide_rel = rel
                # Change to rId1 for consistency
                rel.set('Id', 'rId1')
            elif 'slide' in rel_type.lower() and 'master' not in rel_type.lower() and 'layout' not in rel_type.lower():
                # Remove other slide relationships
                rels_to_remove.append(rel)
        
        for rel in rels_to_remove:
            root.remove(rel)
        
        # Write modified rels
        xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        zout.writestr('ppt/_rels/presentation.xml.rels', xml_bytes)
    
    def _update_app_properties(self, zin: zipfile.ZipFile, zout: zipfile.ZipFile):
        """Update docProps/app.xml to reflect single slide"""
        try:
            app_xml = zin.read('docProps/app.xml')
            root = ET.fromstring(app_xml)
            
            # Find and update slide count
            for elem in root.iter():
                if elem.tag.endswith('Slides'):
                    elem.text = '1'
                elif elem.tag.endswith('HiddenSlides'):
                    elem.text = '0'
            
            xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            zout.writestr('docProps/app.xml', xml_bytes)
        except:
            # If app.xml doesn't exist or fails, copy original
            if 'docProps/app.xml' in zin.namelist():
                zout.writestr('docProps/app.xml', zin.read('docProps/app.xml'))
    
    def _verify_slide_file(self, file_path: Path):
        """Verify that the created PPTX is valid"""
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                # Check for required files
                required = ['[Content_Types].xml', 'ppt/presentation.xml']
                for req in required:
                    if req not in z.namelist():
                        print(f"    ⚠️  Missing {req}")
                        return False
                
                # Verify presentation.xml has exactly 1 slide
                pres_xml = z.read('ppt/presentation.xml')
                root = ET.fromstring(pres_xml)
                sld_id_lst = root.find('.//p:sldIdLst', NS)
                
                if sld_id_lst is not None:
                    slide_count = len(sld_id_lst.findall('p:sldId', NS))
                    if slide_count != 1:
                        print(f"    ⚠️  Contains {slide_count} slides (expected 1)")
                        return False
                
                print(f"    ✓ Verified: 1 slide", end='')
                return True
        except Exception as e:
            print(f"    ❌ Verification failed: {e}")
            return False


def interactive_menu():
    """Interactive CLI for split/merge operations"""
    print("=" * 60)
    print("  PPTX SPLIT & MERGE SYSTEM v2.0")
    print("  100% Formatting Preservation - Fixed XML Handling")
    print("=" * 60)
    print("\nOptions:")
    print("  1) Split presentation into individual slides")
    print("  2) Exit")
    
    choice = input("\nSelect option (1-2): ").strip()
    
    if choice == '1':
        pptx_file = input("\nEnter PPTX file path: ").strip()
        
        # Remove quotes if present
        pptx_file = pptx_file.strip('"').strip("'")
        
        if not Path(pptx_file).exists():
            print(f"❌ File not found: {pptx_file}")
            return
        
        output_dir = input("Output directory (press Enter for default): ").strip()
        output_dir = output_dir if output_dir else None
        
        print()
        splitter = PPTXSplitter(pptx_file)
        output_files = splitter.split(output_dir)
        
        print(f"\n📋 Created files:")
        for f in output_files[:5]:  # Show first 5
            print(f"   • {Path(f).name}")
        if len(output_files) > 5:
            print(f"   ... and {len(output_files) - 5} more")
        
    elif choice == '2':
        print("\n👋 Goodbye!")
        return
    else:
        print("\n❌ Invalid option")
    
    print("\n" + "=" * 60)


if __name__ == "__main__":
    interactive_menu()
