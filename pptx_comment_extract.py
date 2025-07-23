#!/usr/bin/env python3
"""
extract_comments.py ― PPTX コメント抽出スクリプト（PurePosixPath 修正版）

依存: pip install lxml
"""

from pathlib import Path, PurePosixPath
from zipfile import ZipFile
import posixpath
from lxml import etree

PPTX_PATH = Path(__file__).with_name("input.pptx")
NS = {
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def _xml(zf: ZipFile, part: str) -> etree._Element:
    return etree.fromstring(zf.read(part))


def extract_comments_per_slide(pptx: Path) -> dict[int, list[dict]]:
    results: dict[int, list[dict]] = {}

    with ZipFile(pptx) as zf:
        # 1) コメント作者を取得
        author_map = {}
        for part in zf.namelist():
            if part.startswith("ppt/commentAuthors/") and part.endswith(".xml"):
                root = _xml(zf, part)
                for n in root.xpath(".//p:cmAuthor", namespaces=NS):
                    author_map[n.get("id")] = n.get("name")

        # 2) スライドを走査
        slide_parts = sorted(p for p in zf.namelist() if p.startswith("ppt/slides/slide") and p.endswith(".xml"))

        for idx, slide_part in enumerate(slide_parts, start=1):
            rel_part = slide_part.replace("/slides/", "/slides/_rels/") + ".rels"
            if rel_part not in zf.namelist():
                continue

            rel_root = _xml(zf, rel_part)
            rel = rel_root.find(
                './/rel:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"]',
                namespaces=NS,
            )
            if rel is None:
                continue

            # --- ここを修正 ---
            base_dir = PurePosixPath(slide_part).parent  # ppt/slides
            target = PurePosixPath(rel.get("Target"))  # ../comments/commentX.xml
            cm_part = posixpath.normpath(str(base_dir.joinpath(target)))
            # -----------------

            if cm_part not in zf.namelist():
                continue

            cm_root = _xml(zf, cm_part)
            comments = [
                {
                    "author": author_map.get(cm.get("authorId"), "Unknown"),
                    "dt": cm.get("dt"),
                    "text": (cm.find(".//p:text", namespaces=NS).text or "").strip(),
                }
                for cm in cm_root.xpath(".//p:cm", namespaces=NS)
            ]
            if comments:
                results[idx] = comments
    return results


def main() -> None:
    if not PPTX_PATH.exists():
        print(f"❌ {PPTX_PATH} not found")
        return

    data = extract_comments_per_slide(PPTX_PATH)
    if not data:
        print("コメントは見つかりませんでした。")
        return

    for slide_idx in sorted(data):
        print(f"\n=== Slide {slide_idx} ===")
        for c in data[slide_idx]:
            print(f"[{c['dt']}] {c['author']}: {c['text']}")


if __name__ == "__main__":
    main()
