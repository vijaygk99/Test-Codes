"""
Interactive PPTX split & merge tool (with robust XML handling).

- You run the script and answer prompts (no CLI args).
- Split: one slide per PPTX, preserving shapes, text, images, most formatting.
- Merge: takes a folder of PPTX files and merges them in alphabetical order.
- Uses cautious XML cloning + relationship copying to reduce repair/corruption. [web:24][web:25][web:46]

Requires:
    pip install python-pptx
"""

from pptx import Presentation
from copy import deepcopy
from pathlib import Path
from typing import List
import logging
import sys
import os


# ---------- LOGGING ----------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


# ---------- LOW-LEVEL HELPERS (XML + RELS) ----------

def _get_blank_layout(prs: Presentation):
    """
    Return a 'blank-ish' slide layout from prs.

    Choose the layout with the fewest placeholders. [web:24][web:25]
    """
    counts = [len(layout.placeholders) for layout in prs.slide_layouts]
    min_count = min(counts)
    idx = counts.index(min_count)
    return prs.slide_layouts[idx]


def _copy_slide_rels(source_part, dest_part):
    """
    Copy non-notesSlide relationships from source slide part to dest slide part.

    This is a safer version of patterns discussed in GitHub issues/StackOverflow
    to avoid missing images and many charts while not duplicating notesSlide rels. [web:25][web:46]
    """
    for rel in list(source_part.rels.values()):
        # Skip notes slide rel because dest might not have a notes slide
        if "notesSlide" in rel.reltype:
            continue
        try:
            dest_part.rels._add_relationship(rel.reltype, rel._target, False)
        except Exception as e:
            logger.debug(f"Skipping rel {rel.rId} ({rel.reltype}): {e}")


def clone_slide_safe(source_slide, dest_prs: Presentation):
    """
    Clone a slide into dest_prs with XML + relationship handling.

    Steps:
    - Create a new slide using a neutral layout.
    - Delete default placeholders.
    - Deep-copy each shape's XML into the new slide.
    - Copy slide-level relationships (images, charts, etc., excluding notesSlide). [web:25][web:42]
    - Try to copy simple background color.

    This is still limited by python-pptx internals and may not cover 100% of edge cases
    (very complex charts, embedded OLE, etc.), but is a robust general solution.
    """
    # 1) add new slide
    blank_layout = _get_blank_layout(dest_prs)
    new_slide = dest_prs.slides.add_slide(blank_layout)

    # 2) copy relationships first so targets exist when shapes reference them
    try:
        _copy_slide_rels(source_slide.part, new_slide.part)
    except Exception as e:
        logger.debug(f"Could not copy slide relationships: {e}")

    # 3) remove default shapes
    for shp in list(new_slide.shapes):
        el = shp.element
        el.getparent().remove(el)

    # 4) copy shape XML
    from pptx.shapes.shapetree import _SlideShapeTree  # type: ignore

    sp_tree = new_slide.shapes  # _SlideShapeTree
    for shape in source_slide.shapes:
        try:
            new_el = deepcopy(shape.element)
            # insert before extLst (standard pattern) [web:24][web:25]
            sp_tree._spTree.insert_element_before(new_el, "p:extLst")
        except Exception as e:
            logger.debug(f"Error copying shape: {e}")

    # 5) simple background copy
    try:
        src_bg = source_slide.background
        dst_bg = new_slide.background
        if hasattr(src_bg, "fill") and src_bg.fill.type is not None:
            dst_bg.fill.solid()
            if getattr(src_bg.fill.fore_color, "rgb", None):
                dst_bg.fill.fore_color.rgb = src_bg.fill.fore_color.rgb
    except Exception as e:
        logger.debug(f"Error copying background: {e}")

    return new_slide


# ---------- HIGH-LEVEL OPERATIONS ----------

def split_pptx(input_file: str, output_dir: str) -> List[str]:
    """
    Split input_file into N files, each with one slide.
    """
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    if input_path.suffix.lower() != ".pptx":
        raise ValueError("Input must be a .pptx file")

    prs = Presentation(strThis can be made interactive (no CLI args) and it already does XML‑level cloning of shapes and slide relations as far as python‑pptx safely allows, but no solution can literally handle “all scenarios” (especially complex charts, embedded objects, and all master/layout combinations) without hitting library limits. [web:24][web:42][web:25]

Below is a **single interactive script** that:

- Prompts you for:
  - Input PPTX path → splits into one‑slide files.
  - Folder path of split PPTX files → merges them back.
- Uses a **careful XML cloning pattern**:
  - Deep‑copies shape XML into the new slide’s shape tree. [web:24][web:43]
  - Duplicates non‑notes relationships for the slide part (images, charts, media) to reduce chart/embedded‑object breakage. [web:26][web:42][web:25]
- Avoids touching masters/layout parts directly (this is what usually causes repair & duplicate‑name bugs). [web:35][web:50]

```python
"""
Interactive PPTX Split & Merge Tool (no CLI args)

- You run the script.
- It asks for paths.
- Internally it clones slide XML and relationships in a safe way using python-pptx.

Install dependency first:
    pip install python-pptx
"""

from pptx import Presentation
from copy import deepcopy
from pathlib import Path
from typing import List
import logging
import os
import sys

# ---------- LOGGING ----------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


# ---------- LOW-LEVEL SAFE HELPERS ----------

def _get_blank_layout(prs: Presentation):
    """
    Return a 'blank-ish' layout from prs.

    Uses the layout with the fewest placeholders; this is a common safe pattern.[5][1]
    """
    counts = [len(layout.placeholders) for layout in prs.slide_layouts]
    min_count = min(counts)
    idx = counts.index(min_count)
    return prs.slide_layouts[idx]


def _copy_slide_rels(source_slide, dest_slide):
    """
    Copy relationships from source slide part to dest slide part, except notesSlide.

    This helps keep images, media, and most chart parts wired up while avoiding
    known corruption around notes/duplicate rels.[2][3][5]
    """
    try:
        rels = source_slide.part.rels
        for r_id, rel in list(rels.items()):
            if "notesSlide" in rel.reltype:
                continue
            # Use _add_relationship to avoid reusing rId; let pptx generate new one.
            dest_slide.part.rels._add_relationship(
                rel.reltype,
                rel._target,
                is_external=rel.is_external,
            )
    except Exception as e:
        logger.debug(f"Could not copy slide relationships: {e}")


def clone_slide_into_presentation(source_slide, dest_prs: Presentation):
    """
    Clone source_slide into dest_prs:

    - Adds new slide with a neutral layout.
    - Clears default placeholders.
    - Deep-copies each shape's XML _element into the new slide's shape tree.[7][1]
    - Copies simple background color when possible.
    - Copies slide-level relationships for images/media/charts (except notes).[3][2][5]

    Does NOT:
    - Duplicate masters or layout parts (python-pptx does not support this safely).[6][8]
    """
    # Create new slide
    blank_layout = _get_blank_layout(dest_prs)
    new_slide = dest_prs.slides.add_slide(blank_layout)

    # Remove existing placeholders
    for shape in list(new_slide.shapes):
        el = shape.element
        el.getparent().remove(el)

    # Copy shapes XML
    for shape in source_slide.shapes:
        new_el = deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # Copy simple background
    try:
        src_bg = source_slide.background
        dst_bg = new_slide.background
        if hasattr(src_bg, "fill") and src_bg.fill.type is not None:
            dst_bg.fill.solid()
            if getattr(src_bg.fill.fore_color, "rgb", None):
                dst_bg.fill.fore_color.rgb = src_bg.fill.fore_color.rgb
    except Exception as e:
        logger.debug(f"Could not copy background: {e}")

    # Copy relationships (images, media, charts, etc., excluding notes)[2][3]
    _copy_slide_rels(source_slide, new_slide)

    return new_slide


# ---------- CORE OPERATIONS ----------

def split_pptx_one_slide_per_file(input_file: str, output_dir: str) -> List[str]:
    """
    Split a PPTX into multiple PPTX files, each containing a single slide.
    """
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    if input_path.suffix.lower() != ".pptx":
        raise ValueError("Input must be a .pptx file")

    prs = Presentation(str(input_path))
    total_slides = len(prs.slides)
    logger.info("Loaded %s (%d slides)", input_path.name, total_slides)

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    basename = input_path.stem
    output_files: List[str] = []

    for idx, slide in enumerate(prs.slides, start=1):
        new_prs = Presentation()
        new_prs.slide_width = prs.slide_width
        new_prs.slide_height = prs.slide_height

        clone_slide_into_presentation(slide, new_prs)

        out_name = out_dir / f"{basename}_slide_{idx:03d}.pptx"
        new_prs.save(str(out_name))
        output_files.append(str(out_name))
        logger.info("Created split file: %s", out_name)

    logger.info("Split complete: %d files created in %s", len(output_files), out_dir)
    return output_files


def merge_pptx_files(input_files: List[str], output_file: str) -> str:
    """
    Merge a list of PPTX files (in order) into a single PPTX.
    """
    if not input_files:
        raise ValueError("No input files provided for merge")

    paths: List[Path] = []
    for f in input_files:
        p = Path(f)
        if not p.exists():
            raise FileNotFoundError(f"File not found: {f}")
        if p.suffix.lower() != ".pptx":
            raise ValueError(f"Not a .pptx file: {f}")
        paths.append(p)

    base_prs = Presentation(str(paths))
    logger.info("Base presentation: %s (%d slides)", paths.name, len(base_prs.slides))

    for p in paths[1:]:
        logger.info("Merging from: %s", p.name)
        src = Presentation(str(p))
        for slide in src.slides:
            clone_slide_into_presentation(slide, base_prs)

    out_path = Path(output_file)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    base_prs.save(str(out_path))
    logger.info("Merge complete: %s (%d slides)", out_path, len(base_prs.slides))

    return str(out_path)


def merge_pptx_from_directory(input_dir: str, output_file: str) -> str:
    """
    Merge all PPTX files from a directory (sorted by name) into one PPTX.
    """
    dir_path = Path(input_dir)
    if not dir_path.exists():
        raise FileNotFoundError(f"Directory not found: {input_dir}")

    pptx_paths = sorted(
        [p for p in dir_path.iterdir() if p.is_file() and p.suffix.lower() == ".pptx"],
        key=lambda p: p.name,
    )
    if not pptx_paths:
        raise FileNotFoundError(f"No .pptx files found in directory: {input_dir}")

    return merge_pptx_files([str(p) for p in pptx_paths], output_file)


# ---------- INTERACTIVE PROMPT UI ----------

def prompt_path(prompt: str, must_exist: bool = True, is_dir: bool = False) -> str:
    while True:
        val = input(prompt).strip().strip('"').strip("'")
        if not val:
            print("Path cannot be empty.")
            continue
        p = Path(val)
        if must_exist:
            if not p.exists():
                print(f"Path does not exist: {p}")
                continue
            if is_dir and not p.is_dir():
                print(f"Not a directory: {p}")
                continue
            if not is_dir and p.is_dir():
                print(f"Expected a file, got a directory: {p}")
                continue
        return str(p)


def main_menu():
    print("\n" + "=" * 60)
    print(" PPTX SPLIT & MERGE (INTERACTIVE)")
    print("=" * 60)
    print("1) Split a PPTX into one-slide files")
    print("2) Merge PPTX files from a directory (sorted)")
    print("3) Exit")
    print("-" * 60)


def run():
    while True:
        main_menu()
        choice = input("Choose option (1-3): ").strip()
        if choice == "1":
            print("\n--- SPLIT MODE ---")
            input_file = prompt_path("Enter path to input PPTX file: ", must_exist=True, is_dir=False)
            default_out = str(Path(input_file).with_suffix("")) + "_split"
            out_dir = input(f"Enter output folder (default: {default_out}): ").strip()
            if not out_dir:
                out_dir = default_out
            try:
                files = split_pptx_one_slide_per_file(input_file, out_dir)
                print(f"\nSplit done. {len(files)} files written to: {out_dir}")
            except Exception as e:
                print(f"Error during split: {e}")
                logger.error("Split error", exc_info=True)

        elif choice == "2":
            print("\n--- MERGE MODE ---")
            in_dir = prompt_path("Enter directory with PPTX files to merge: ", must_exist=True, is_dir=True)
            default_out = str(Path(in_dir).with_name("merged_output.pptx"))
            out_file = input(f"Enter output PPTX path (default: {default_out}): ").strip()
            if not out_file:
                out_file = default_out
            try:
                merged = merge_pptx_from_directory(in_dir, out_file)
                print(f"\nMerge done. Output file: {merged}")
            except Exception as e:
                print(f"Error during merge: {e}")
                logger.error("Merge error", exc_info=True)

        elif choice == "3":
            print("\nExiting. Bye.")
            break
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")


if __name__ == "__main__":
    try:
        run()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(0)
