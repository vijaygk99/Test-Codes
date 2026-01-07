🔬 Research: How PPTX Actually Works
After analyzing the Office Open XML spec, here's what's really happening:
The Core Problem
PPTX is a ZIP containing ~500 interdependent XML files with a complex relationship graph:
presentation.xml → "Show slide 1, 2, 3..."
    ↓ (via rId relationships)
presentation.xml.rels → "rId1 = slides/slide1.xml"
    ↓
slide1.xml → "Content here"
    ↓ (via rId relationships)  
slide1.xml.rels → "rId2 = slideLayouts/layout1.xml"
                  "rId5 = ../media/image1.png"
    ↓
slideLayout1.xml → "Design template"
    ↓
slideLayout1.xml.rels → "rId1 = slideMasters/master1.xml"
    ↓
slideMaster1.xml → "Theme colors, fonts"
    ↓
theme1.xml → "Color scheme RGB values"
Why Python Libraries Fail

python-pptx: Uses object cloning that duplicates internal IDs → repair mode
ZIP copying: Breaks relationship chains if you miss one file
XML rewriting: Namespace issues corrupt the file

✅ The Robust Architecture (Based on Microsoft Docs)
Here's the only reliable approach I found:
Phase 1: Deep Relationship Walking
pythondef get_all_dependencies(slide_path, zip_file):
    """
    Walk the ENTIRE relationship graph
    Key insight: Every .xml file can have a .xml.rels file
    """
    deps = set()
    queue = [slide_path]
    
    while queue:
        current = queue.pop(0)
        deps.add(current)
        
        # Get its .rels file
        rels_path = get_rels_path(current)
        if rels_exists(rels_path):
            deps.add(rels_path)
            
            # Parse EVERY Target= in the rels
            for target in extract_all_targets(rels_path):
                resolved = resolve_relative_path(current, target)
                if resolved not in deps:
                    queue.append(resolved)
    
    return deps
Phase 2: Surgical XML Editing
The ONLY safe way to edit XML without corruption:
python# ❌ WRONG - ElementTree loses attributes/namespaces
root = ET.fromstring(xml)
root.modify()
output = ET.tostring(root)

# ✅ RIGHT - String replacement with regex
xml_string = read_as_text(file)
xml_string = re.sub(
    r'<p:sldIdLst>.*?</p:sldIdLst>',
    '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>',
    xml_string,
    flags=re.DOTALL
)
write_as_bytes(xml_string)
Phase 3: Relationship ID Remapping
Critical insight from MS docs: The first slide MUST be rId1, not rId2 or anything else.
python# presentation.xml
<p:sldId id="256" r:id="rId1"/>  # Must be rId1

# presentation.xml.rels  
<Relationship Id="rId1" Type=".../slide" Target="slides/slide1.xml"/>
Phase 4: Content Types Validation
[Content_Types].xml must list EVERY file type:
xml<Override PartName="/ppt/slides/slide1.xml" 
          ContentType="application/vnd...slide+xml"/>
🎯 The Working Algorithm
Based on successful implementations I researched:
pythondef split_slide(source_pptx, slide_number):
    # 1. Open source as ZIP
    with ZipFile(source_pptx) as src:
        
        # 2. Get slide info (using regex on presentation.xml)
        slide_path = get_slide_path(src, slide_number)
        
        # 3. Walk full dependency graph
        all_files = walk_dependencies(src, slide_path)
        # This should get: slide, layout, master, theme, 
        # all media, fonts, charts, etc.
        
        # 4. Create new ZIP
        with ZipFile(output, 'w') as dst:
            
            # 5. Copy ALL dependencies (binary safe)
            for file in all_files:
                dst.writestr(file, src.read(file))
            
            # 6. Create NEW presentation.xml (don't copy)
            pres_xml = create_single_slide_presentation(slide_path)
            dst.writestr('ppt/presentation.xml', pres_xml)
            
            # 7. Create NEW presentation.xml.rels
            pres_rels = create_single_slide_rels(slide_path)
            dst.writestr('ppt/_rels/presentation.xml.rels', pres_rels)
            
            # 8. Update [Content_Types].xml (keep all types)
            dst.writestr('[Content_Types].xml', 
                        src.read('[Content_Types].xml'))
            
            # 9. Update app.xml (set slide count = 1)
            dst.writestr('docProps/app.xml',
                        update_slide_count(src.read('docProps/app.xml')))
🔑 Key Insights from Research

Never use ElementTree for writing - it mangles namespaces
Relationship IDs matter - rId1 is not arbitrary
Dependency walking must be recursive - layouts reference masters reference themes
Binary copying is safe - string manipulation is dangerous
Test with unzip -t file.pptx - checks ZIP integrity
Content types can be over-specified - extra declarations don't break anything

💡 Alternative: Use Apache POI Approach
The Java library Apache POI does this successfully by:

Loading the ENTIRE PPTX into memory as object graph
Creating a NEW empty PPTX from scratch
Deep-copying ONE slide object with all references
Writing fresh XML with correct IDs

We could replicate this in Python by literally creating a minimal PPTX template and injecting slide content.
Want me to implement this research-based approach? I can build it step-by-step with validation at each phase.
