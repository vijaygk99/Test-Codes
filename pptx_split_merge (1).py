#!/usr/bin/env python3
"""
PPTX Split System v4.0 - Complete Rebuild
Handles all PPTX structures with proper namespace handling
Tests each step and validates output
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Set, List, Tuple
import re
from io import BytesIO

# Register all PPTX namespaces
NAMESPACES = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'
}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

class PPTXSplitter:
    """Splits PPTX preserving all formatting and structure"""
    
    def __init__(self, input_pptx: str):
        self.input_path = Path(input_pptx)
        self.base_name = self.input_path.stem
        
    def split(self, output_dir: str = None) -> List[str]:
        """Split PPTX into individual slide files"""
        if output_dir is None:
            output_dir = f"{self.base_name}_slides"
        
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        print(f"\n{'='*70}")
        print(f"📂 Opening: {self.input_path.name}")
        print(f"{'='*70}\n")
        
        try:
            with zipfile.ZipFile(self.input_path, 'r') as source_zip:
                
                # Step 1: Analyze the PPTX structure
                print("🔍 STEP 1: Analyzing PPTX Structure...")
                analysis = self._analyze_pptx(source_zip)
                
                if not analysis['valid']:
                    print(f"❌ Invalid PPTX: {analysis['error']}")
                    return []
                
                print(f"   ✓ Found {analysis['slide_count']} slides")
                print(f"   ✓ Found {len(analysis['all_files'])} total files in ZIP")
                print()
                
                # Step 2: Extract slide information
                print("🔍 STEP 2: Reading Slide Information...")
                slides = self._get_slides_with_details(source_zip)
                
                if not slides:
                    print("❌ No slides found in presentation!")
                    return []
                
                print(f"   ✓ Successfully read {len(slides)} slide(s):")
                for idx, slide in enumerate(slides, 1):
                    print(f"      Slide {idx}: {slide['path']}")
                print()
                
                # Step 3: Create individual slide files
                print("🔍 STEP 3: Creating Individual Slide Files...")
                output_files = []
                
                for idx, slide_data in enumerate(slides, 1):
                    output_file = output_path / f"{self.base_name}_slide_{idx:03d}.pptx"
                    print(f"\n   📄 Creating: {output_file.name}")
                    
                    success = self._create_single_slide_file(
                        source_zip, 
                        slide_data, 
                        output_file, 
                        idx
                    )
                    
                    if success:
                        output_files.append(str(output_file))
                        print(f"      ✅ Success - Slide {idx} created")
                    else:
                        print(f"      ❌ Failed to create slide {idx}")
                
                print(f"\n{'='*70}")
                print(f"✨ COMPLETE: Created {len(output_files)}/{len(slides)} slides")
                print(f"📁 Output: {output_path}/")
                print(f"{'='*70}\n")
                
                return output_files
                
        except Exception as e:
            print(f"\n❌ FATAL ERROR: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _analyze_pptx(self, source_zip: zipfile.ZipFile) -> Dict:
        """Analyze PPTX structure and validate"""
        try:
            all_files = source_zip.namelist()
            
            # Check for required files
            required = ['ppt/presentation.xml', '[Content_Types].xml']
            for req in required:
                if req not in all_files:
                    return {'valid': False, 'error': f'Missing required file: {req}'}
            
            # Read and parse presentation.xml
            pres_xml = source_zip.read('ppt/presentation.xml')
            pres_root = ET.fromstring(pres_xml)
            
            # Count slides
            slide_count = 0
            sld_id_lst = pres_root.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst')
            
            if sld_id_lst is not None:
                slide_count = len(list(sld_id_lst))
            
            return {
                'valid': True,
                'slide_count': slide_count,
                'all_files': all_files,
                'has_rels': 'ppt/_rels/presentation.xml.rels' in all_files
            }
            
        except Exception as e:
            return {'valid': False, 'error': str(e)}
    
    def _get_slides_with_details(self, source_zip: zipfile.ZipFile) -> List[Dict]:
        """Extract complete slide information with all metadata"""
        slides = []
        
        try:
            # Parse presentation.xml
            pres_xml = source_zip.read('ppt/presentation.xml')
            pres_root = ET.fromstring(pres_xml)
            
            # Parse presentation rels
            pres_rels_xml = source_zip.read('ppt/_rels/presentation.xml.rels')
            pres_rels_root = ET.fromstring(pres_rels_xml)
            
            # Find all slides
            sld_id_lst = pres_root.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst')
            
            if sld_id_lst is None:
                return []
            
            # Process each slide
            for sld_id_elem in sld_id_lst:
                rel_id = sld_id_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                slide_id = sld_id_elem.get('id')
                
                if not rel_id:
                    continue
                
                # Find the corresponding relationship
                for rel in pres_rels_root:
                    if rel.get('Id') == rel_id:
                        target = rel.get('Target')
                        if target:
                            slide_path = f"ppt/{target}"
                            
                            slides.append({
                                'rel_id': rel_id,
                                'slide_id': slide_id,
                                'path': slide_path,
                                'target': target
                            })
                            break
            
            return slides
            
        except Exception as e:
            print(f"   ❌ Error reading slides: {e}")
            return []
    
    def _create_single_slide_file(self, source_zip: zipfile.ZipFile, 
                                  slide_data: Dict, output_file: Path, 
                                  slide_num: int) -> bool:
        """Create a complete valid PPTX with one slide"""
        try:
            # Collect all dependencies
            print(f"      → Collecting dependencies...")
            all_deps = self._collect_all_dependencies(source_zip, slide_data)
            print(f"      → Found {len(all_deps)} dependent files")
            
            # Create the output PPTX
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as out_zip:
                
                # Copy all dependency files (binary safe)
                print(f"      → Copying {len(all_deps)} files...")
                for file_path in all_deps:
                    if file_path in ['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels']:
                        continue  # We'll write these specially
                    
                    if file_path in source_zip.namelist():
                        out_zip.writestr(file_path, source_zip.read(file_path))
                
                # Write modified presentation.xml (with only this slide)
                print(f"      → Writing presentation.xml...")
                self._write_modified_presentation(source_zip, out_zip, slide_data)
                
                # Write modified presentation.xml.rels (with only this slide)
                print(f"      → Writing presentation rels...")
                self._write_modified_presentation_rels(source_zip, out_zip, slide_data)
                
                # Update app.xml
                print(f"      → Updating properties...")
                self._update_app_properties(source_zip, out_zip)
            
            # Verify the created file
            print(f"      → Verifying...")
            return self._verify_output(output_file)
            
        except Exception as e:
            print(f"      ❌ Error: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _collect_all_dependencies(self, source_zip: zipfile.ZipFile, 
                                  slide_data: Dict) -> Set[str]:
        """Collect ALL files needed for this slide"""
        deps = set()
        to_process = []
        processed = set()
        
        # Start with the slide
        slide_path = slide_data['path']
        to_process.append(slide_path)
        
        # Process dependency chain
        while to_process:
            current = to_process.pop(0)
            
            if current in processed:
                continue
            
            processed.add(current)
            deps.add(current)
            
            # Add the .rels file for this item
            rels_path = self._get_rels_path(current)
            if rels_path in source_zip.namelist():
                deps.add(rels_path)
                
                # Parse rels to find dependencies
                try:
                    rels_xml = source_zip.read(rels_path)
                    rels_root = ET.fromstring(rels_xml)
                    
                    for rel in rels_root:
                        target = rel.get('Target')
                        if target and not target.startswith('http'):
                            # Resolve relative path
                            dep_path = self._resolve_path(current, target)
                            if dep_path and dep_path not in processed:
                                to_process.append(dep_path)
                except:
                    pass
        
        # Add core structure files
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
            if core in source_zip.namelist():
                deps.add(core)
        
        # Add all theme files
        for fname in source_zip.namelist():
            if '/theme' in fname:
                deps.add(fname)
        
        return deps
    
    def _get_rels_path(self, file_path: str) -> str:
        """Get relationship file path"""
        parts = file_path.rsplit('/', 1)
        if len(parts) == 2:
            return f"{parts[0]}/_rels/{parts[1]}.rels"
        return f"_rels/{file_path}.rels"
    
    def _resolve_path(self, base: str, target: str) -> str:
        """Resolve relative path"""
        if target.startswith('/'):
            return target[1:]
        
        base_dir = base.rsplit('/', 1)[0] if '/' in base else ''
        parts = target.split('/')
        base_parts = base_dir.split('/') if base_dir else []
        
        for part in parts:
            if part == '..':
                if base_parts:
                    base_parts.pop()
            elif part and part != '.':
                base_parts.append(part)
        
        return '/'.join(base_parts) if base_parts else None
    
    def _write_modified_presentation(self, source_zip: zipfile.ZipFile,
                                    out_zip: zipfile.ZipFile, slide_data: Dict):
        """Write presentation.xml with only this slide"""
        
        # Read original
        pres_xml = source_zip.read('ppt/presentation.xml')
        pres_root = ET.fromstring(pres_xml)
        
        # Find sldIdLst
        sld_id_lst = pres_root.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst')
        
        if sld_id_lst is not None:
            # Remove all children
            for child in list(sld_id_lst):
                sld_id_lst.remove(child)
            
            # Add only our slide with new ID
            new_slide = ET.Element(
                '{http://schemas.openxmlformats.org/presentationml/2006/main}sldId',
                attrib={
                    'id': '256',
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': 'rId1'
                }
            )
            sld_id_lst.append(new_slide)
        
        # Write to output
        xml_bytes = ET.tostring(pres_root, encoding='utf-8', xml_declaration=True)
        out_zip.writestr('ppt/presentation.xml', xml_bytes)
    
    def _write_modified_presentation_rels(self, source_zip: zipfile.ZipFile,
                                         out_zip: zipfile.ZipFile, slide_data: Dict):
        """Write presentation.xml.rels with only necessary relationships"""
        
        # Read original
        rels_xml = source_zip.read('ppt/_rels/presentation.xml.rels')
        rels_root = ET.fromstring(rels_xml)
        
        # Remove all slide relationships
        for rel in list(rels_root):
            rel_type = rel.get('Type', '')
            if '/slide' in rel_type and 'slideMaster' not in rel_type and 'slideLayout' not in rel_type:
                rels_root.remove(rel)
        
        # Add our slide as rId1
        new_rel = ET.Element(
            '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',
            attrib={
                'Id': 'rId1',
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                'Target': slide_data['target']
            }
        )
        rels_root.append(new_rel)
        
        # Write to output
        xml_bytes = ET.tostring(rels_root, encoding='utf-8', xml_declaration=True)
        out_zip.writestr('ppt/_rels/presentation.xml.rels', xml_bytes)
    
    def _update_app_properties(self, source_zip: zipfile.ZipFile, out_zip: zipfile.ZipFile):
        """Update app.xml to show 1 slide"""
        try:
            if 'docProps/app.xml' in source_zip.namelist():
                app_xml = source_zip.read('docProps/app.xml')
                app_str = app_xml.decode('utf-8')
                
                # Update counts
                app_str = re.sub(r'<Slides>\d+</Slides>', '<Slides>1</Slides>', app_str)
                app_str = re.sub(r'<HiddenSlides>\d+</HiddenSlides>', '<HiddenSlides>0</HiddenSlides>', app_str)
                
                out_zip.writestr('docProps/app.xml', app_str.encode('utf-8'))
        except:
            pass
    
    def _verify_output(self, output_file: Path) -> bool:
        """Verify the output file is valid"""
        try:
            with zipfile.ZipFile(output_file, 'r') as z:
                # Check presentation.xml
                pres_xml = z.read('ppt/presentation.xml')
                pres_root = ET.fromstring(pres_xml)
                
                # Count slides
                sld_id_lst = pres_root.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst')
                if sld_id_lst is not None:
                    slide_count = len(list(sld_id_lst))
                    if slide_count == 1:
                        return True
                    else:
                        print(f"      ⚠️  Expected 1 slide, found {slide_count}")
                        return False
            
            return True
        except Exception as e:
            print(f"      ⚠️  Verification error: {e}")
            return False


def main():
    """Main entry point"""
    print("\n" + "="*70)
    print("  PPTX SPLITTER v4.0")
    print("  Step-by-step validation with detailed logging")
    print("="*70)
    
    pptx_file = input("\n📎 Enter PPTX file path: ").strip().strip('"').strip("'")
    
    if not Path(pptx_file).exists():
        print(f"\n❌ File not found: {pptx_file}\n")
        return
    
    output_dir = input("📁 Output directory (press Enter for default): ").strip()
    output_dir = output_dir if output_dir else None
    
    splitter = PPTXSplitter(pptx_file)
    output_files = splitter.split(output_dir)
    
    if output_files:
        print("\n✅ SUCCESS! Test by opening any of the created files.")
    else:
        print("\n❌ FAILED! Check the error messages above.")
    
    print()


if __name__ == "__main__":
    main()
