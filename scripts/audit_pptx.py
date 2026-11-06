#!/usr/bin/env python3
"""Report presentation geometry and readability warnings without editing the file."""

import argparse
import json
from collections import Counter
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def rect(shape):
    return (shape.left, shape.top, shape.left + shape.width, shape.top + shape.height)


def overlap(a, b):
    left = max(a[0], b[0])
    top = max(a[1], b[1])
    right = min(a[2], b[2])
    bottom = min(a[3], b[3])
    if right <= left or bottom <= top:
        return 0
    return (right - left) * (bottom - top)


def area(box):
    return max(0, box[2] - box[0]) * max(0, box[3] - box[1])


def text_of(shape):
    if getattr(shape, "has_text_frame", False):
        return shape.text.strip()
    return ""


def audit(path, expected_pages, min_font):
    prs = Presentation(path)
    report = {
        "file": str(path),
        "slide_count": len(prs.slides),
        "slide_width_inches": round(prs.slide_width / 914400, 3),
        "slide_height_inches": round(prs.slide_height / 914400, 3),
        "expected_pages": expected_pages,
        "page_count_matches": expected_pages is None or len(prs.slides) == expected_pages,
        "small_text_threshold_pt": min_font,
        "warnings": [],
        "slides": [],
    }
    fonts = Counter()
    sizes = Counter()

    if expected_pages is not None and len(prs.slides) != expected_pages:
        report["warnings"].append(f"Expected {expected_pages} slides, found {len(prs.slides)}")
    ratio = prs.slide_width / prs.slide_height
    if abs(ratio - 16 / 9) > 0.01:
        report["warnings"].append(f"Slide ratio is {ratio:.4f}, not 16:9")

    for index, slide in enumerate(prs.slides, start=1):
        slide_info = {
            "slide": index,
            "pictures": 0,
            "tables": 0,
            "small_text": [],
            "out_of_bounds": [],
            "picture_overlaps": [],
            "text_overlaps": [],
        }
        pictures = []
        texts = []
        for shape_index, shape in enumerate(slide.shapes, start=1):
            box = rect(shape)
            name = shape.name or f"Shape {shape_index}"
            if shape.left < 0 or shape.top < 0 or box[2] > prs.slide_width or box[3] > prs.slide_height:
                slide_info["out_of_bounds"].append({"shape": name, "box_emu": box})
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                slide_info["pictures"] += 1
                pictures.append((name, box))
            if getattr(shape, "has_table", False):
                slide_info["tables"] += 1
            text = text_of(shape)
            if text:
                texts.append((name, box, text[:100]))
            if getattr(shape, "has_text_frame", False):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        content = run.text.strip()
                        if not content:
                            continue
                        if run.font.name:
                            fonts[run.font.name] += len(content)
                        if run.font.size:
                            pt = round(run.font.size.pt, 1)
                            sizes[pt] += len(content)
                            if pt < min_font:
                                slide_info["small_text"].append({
                                    "shape": name,
                                    "size_pt": pt,
                                    "text": content[:100],
                                })

        for i, (name_a, box_a) in enumerate(pictures):
            for name_b, box_b in pictures[i + 1:]:
                common = overlap(box_a, box_b)
                smaller = min(area(box_a), area(box_b))
                if smaller and common / smaller > 0.02:
                    slide_info["picture_overlaps"].append({
                        "shape_a": name_a,
                        "shape_b": name_b,
                        "smaller_picture_overlap_ratio": round(common / smaller, 3),
                    })

        for i, (name_a, box_a, text_a) in enumerate(texts):
            for name_b, box_b, text_b in texts[i + 1:]:
                common = overlap(box_a, box_b)
                smaller = min(area(box_a), area(box_b))
                if smaller and common / smaller > 0.20:
                    slide_info["text_overlaps"].append({
                        "shape_a": name_a,
                        "shape_b": name_b,
                        "text_a": text_a,
                        "text_b": text_b,
                        "smaller_textbox_overlap_ratio": round(common / smaller, 3),
                    })

        report["slides"].append(slide_info)

    report["fonts_by_character_count"] = fonts.most_common()
    report["font_sizes_by_character_count"] = sorted(sizes.items())
    report["warning_counts"] = {
        "small_text": sum(len(s["small_text"]) for s in report["slides"]),
        "out_of_bounds": sum(len(s["out_of_bounds"]) for s in report["slides"]),
        "picture_overlaps": sum(len(s["picture_overlaps"]) for s in report["slides"]),
        "text_overlaps": sum(len(s["text_overlaps"]) for s in report["slides"]),
    }
    return report


def main():
    parser = argparse.ArgumentParser(description="Audit PPTX geometry and text readability")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--pages", type=int, default=None)
    parser.add_argument("--min-font", type=float, default=12.0)
    parser.add_argument("-o", "--output", type=Path)
    args = parser.parse_args()
    source = args.pptx.expanduser().resolve()
    if not source.is_file():
        parser.error(f"File not found: {source}")
    report = audit(source, args.pages, args.min_font)
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        output = args.output.expanduser().resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(payload + "\n", encoding="utf-8")
    print(json.dumps({
        "file": report["file"],
        "slides": report["slide_count"],
        "page_count_matches": report["page_count_matches"],
        "warning_counts": report["warning_counts"],
        "output": str(args.output) if args.output else None,
        "note": "Overlap and small-font findings require visual review; this is not a rendering check.",
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
