"""
Safe, robust PPTX split & merge tool using python-pptx.

- Split: one slide per PPTX (you give only the .pptx path and an output folder).
- Merge: merge any list of PPTX files back into one (you give paths only).
- Uses only safe, public python-pptx APIs to avoid corrupted files and
  'duplicate name' / repair mode issues. [web:24][web:31][web:42]

Install:
    pip install python-pptx

Usage examples (from terminal):

    # Split all slides into separate PPTX files
    python pptx_safe_tool.py split "input.pptx" "split_output"

    # Merge back some or all split files (order matters)
    python pptx_safe_tool.py merge "merged.pptx" split_output/input_slide_001.pptx split_output/input_slide_002.pptx

    # Merge all PPTX files in a folder (alphabetical order)
    python pptx_safe_tool.py merge_dir "merged.pptx" "split_output"
"""

from pptx import Presentation
from copy import deepcopy
from pathlib import Path
from typing import List
import logging
import sys
import os


# ---------------- LOGGING SETUP ----------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


# ---------------- LOW-LEVEL SAFE HELPERS ----------------

def _get_blank_layout(prs: Presentation):
    """
    Get a 'blank-ish' layout from a presentation.

    Strategy: choose slide layout with the fewest placeholders.
    This is a widely used safe pattern to add arbitrary content. [web:24][web:31]
    """
    counts = [len(layout.placeholders) for layout in prs.slide_layouts]
    min_count = min(counts)
    idx = counts.index(min_count)
    return prs.slide_layouts[idx]


def clone_slide_into_presentation(source_slide, dest_prs: Presentation):
    """
    Safely clone a single slide from one presentation into dest_prs.

    - Creates a new slide using a neutral/blank layout.
    - Removes default placeholders.
    - Deep-copies each shape's XML from the source slide to the new slide.
    - Copies simple background color when possible.

    Does NOT:
    - Clone masters/layouts directly (that is unsupported and corrupts files). [web:35][web:42]
    """
    # Create new slide using a 'blank' layout
    blank_layout = _get_blank_layout(dest_prs)
    new_slide = dest_prs.slides.add_slide(blank_layout)

    # Remove any placeholders on the new slide
    for shape in list(new_slide.shapes):
        el = shape.element
        el.getparent().remove(el)

    # Copy shapes from source slide
    for shape in source_slide.shapes:
        new_el = deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # Copy simple background (if possible)
    try:
        src_bg = source_slide.background
        dst_bg = new_slide.background
        if hasattr(src_bg, "fill") and src_bg.fill.type is not None:
            dst_bg.fill.solid()
            if getattr(src_bg.fill.fore_color, "rgb", None):
                dst_bg.fill.fore_color.rgb = src_bg.fill.fore_color.rgb
    except Exception as e:
        logger.debug(f"Could not copy background: {e}")

    return new_slide


# ---------------- SPLIT FUNCTION ----------------

def split_pptx_one_slide_per_file(input_file: str, output_dir: str) -> List[str]:
    """
    Split a PPTX into multiple PPTX files, each containing one slide.

    Args:
        input_file: path to the input .pptx.
        output_dir: directory where split files will be written.

    Returns:
        List of output file paths (one per slide) in order.
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
        # Create a fresh Presentation
        new_prs = Presentation()
        # Match slide size
        new_prs.slide_width = prs.slide_width
        new_prs.slide_height = prs.slide_height

        # Clone this slide into new presentation
        clone_slide_into_presentation(slide, new_prs)

        out_name = out_dir / f"{basename}_slide_{idx:03d}.pptx"
        new_prs.save(str(out_name))
        output_files.append(str(out_name))
        logger.info("Created split file: %s", out_name)

    logger.info("Split complete: %d files created in %s", len(output_files), out_dir)
    return output_files


# ---------------- MERGE FUNCTIONS ----------------

def merge_pptx_files(input_files: List[str], output_file: str) -> str:
    """
    Merge multiple PPTX files into a single PPTX.

    Args:
        input_files: list of .pptx paths, in desired order.
        output_file: output .pptx path.

    Returns:
        Path to merged PPTX.
    """
    if not input_files:
        raise ValueError("No input files provided for merge")

    pptx_paths: List[Path] = []
    for f in input_files:
        p = Path(f)
        if not p.exists():
            raise FileNotFoundError(f"File not found: {f}")
        if p.suffix.lower() != ".pptx":
            raise ValueError(f"Not a .pptx file: {f}")
        pptx_paths.append(p)

    # Start with first file as base
    base_prs = Presentation(str(pptx_paths[0]))
    logger.info("Base presentation: %s (%d slides)", pptx_paths[0].name, len(base_prs.slides))

    # Append slides from subsequent files
    for p in pptx_paths[1:]:
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
    Convenience: merge all .pptx files in a directory (sorted by name).

    Args:
        input_dir: directory containing .pptx files.
        output_file: output .pptx path.

    Returns:
        Path to merged PPTX.
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

    logger.info("Found %d PPTX files in %s", len(pptx_paths), dir_path)
    return merge_pptx_files([str(p) for p in pptx_paths], output_file)


# ---------------- SIMPLE CLI INTERFACE ----------------

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Safe PPTX split & merge tool (paths only, XML handled internally)."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # split
    split_p = subparsers.add_parser("split", help="Split PPTX into one-slide files")
    split_p.add_argument("input", help="Input .pptx file path")
    split_p.add_argument("output_dir", help="Output directory for split files")

    # merge explicit list
    merge_p = subparsers.add_parser("merge", help="Merge specified PPTX files")
    merge_p.add_argument("output", help="Output .pptx file path")
    merge_p.add_argument("inputs", nargs="+", help="Input .pptx files (in order)")

    # merge all from directory
    merge_dir_p = subparsers.add_parser(
        "merge_dir", help="Merge all .pptx files from a directory (sorted by name)"
    )
    merge_dir_p.add_argument("output", help="Output .pptx file path")
    merge_dir_p.add_argument("input_dir", help="Directory containing .pptx files")

    args = parser.parse_args()

    try:
        if args.command == "split":
            split_pptx_one_slide_per_file(args.input, args.output_dir)
        elif args.command == "merge":
            merge_pptx_files(args.inputs, args.output)
        elif args.command == "merge_dir":
            merge_pptx_from_directory(args.input_dir, args.output)
    except Exception as e:
        logger.error("Error: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
