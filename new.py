#!/usr/bin/env python3
"""
pptx_split_animation.py
~~~~~~~~~~~~~~~~~~~~~~~

End‑to‑end utility: read *input.pptx*, expand every animated slide
(*any* style.visibility toggle) into a sequence of static slides,
and save as *output.pptx*.

変更点
------
1. collect_visibility_events
   • 最近傍の <p:cTn> をキーに “同時刻グループ” を採番
2. build_snapshots
   • (step_serial, spid, visible) に対応し，同一 step_serial を 1 枚に統合

他のロジック・制約・依存ライブラリは従来と同じです。
"""

import zipfile
import shutil
import os
import tempfile
import copy
import re
from pathlib import Path
from lxml import etree as ET

INPUT_PPTX = "input_2.pptx"
OUTPUT_PPTX = "output_2.pptx"

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
}


# --------------------------------------------------------------------------- #
# Utility XML functions                                                       #
# --------------------------------------------------------------------------- #
def next_numeric_id(existing, prefix="rId"):
    nums = [int(re.sub(r"\D", "", s)) for s in existing if re.sub(r"\D", "", s)]
    return f"{prefix}{max(nums, default=0) + 1}"


def collect_shapes(slide_tree):
    shapes = {}
    for el in slide_tree.xpath("//p:sp | //p:pic | //p:graphicFrame", namespaces=NS):
        cNvPr = el.find(".//p:cNvPr", namespaces=NS)
        if cNvPr is None:
            continue
        try:
            spid = int(cNvPr.get("id"))
            shapes[spid] = el
        except (TypeError, ValueError):
            continue
    return shapes


def collect_visibility_events(slide_tree):
    """
    Return list of tuples: (step_serial, spid, visible)
    同じ <p:cTn>（＝同時刻）の <p:set> は同じ step_serial を持つ。
    """
    events = []
    ctn_to_serial = {}
    step_serial = -1

    for set_el in slide_tree.xpath("//p:set", namespaces=NS):
        # 1) style.visibility を変える <p:set> だけ対象
        attr_names = set_el.xpath(".//p:attrNameLst/p:attrName/text()", namespaces=NS)
        if "style.visibility" not in attr_names:
            continue

        spid_attr = set_el.xpath(".//p:spTgt/@spid", namespaces=NS)
        to_val = set_el.xpath("./p:to/p:strVal/@val", namespaces=NS)
        if not spid_attr or not to_val:
            continue

        try:
            spid = int(spid_attr[0])
        except ValueError:
            continue
        visible = to_val[0] != "hidden"

        # 2) 最近傍 <p:cTn> をグループキーに
        ctn = set_el.xpath("ancestor::p:cTn[1]", namespaces=NS)
        ctn_key = ctn[0].get("id") if ctn and ctn[0].get("id") else id(ctn[0] if ctn else set_el)

        if ctn_key not in ctn_to_serial:
            step_serial += 1
            ctn_to_serial[ctn_key] = step_serial
        events.append((ctn_to_serial[ctn_key], spid, visible))

    return events


def build_snapshots(shapes, events):
    """
    events: [(step_serial, spid, visible), …]  で昇順に並んでいる前提。
    """
    state = {spid: True for spid in shapes}

    # 出現アニメの初回は最初は非表示に
    first_change = {}
    for _, spid, vis in events:
        if spid not in first_change:
            first_change[spid] = vis
            if vis:
                state[spid] = False

    snapshots = [copy.deepcopy(state)]  # step0（初期）

    current_step = -1
    for step_serial, spid, visible in events:
        if step_serial != current_step:
            if current_step != -1:  # -1 は初期
                snapshots.append(copy.deepcopy(state))
            current_step = step_serial
        state[spid] = visible

    snapshots.append(copy.deepcopy(state))  # 最終状態
    return snapshots


def materialise_snapshot(orig_tree, visible_map):
    tree = copy.deepcopy(orig_tree)
    timing = tree.find(".//p:timing", namespaces=NS)
    if timing is not None:
        timing.getparent().remove(timing)

    for el in tree.xpath("//p:sp | //p:pic | //p:graphicFrame", namespaces=NS):
        cNvPr = el.find(".//p:cNvPr", namespaces=NS)
        if cNvPr is None:
            continue
        try:
            spid = int(cNvPr.get("id"))
        except (TypeError, ValueError):
            continue

        if not visible_map.get(spid, True):
            el.getparent().remove(el)

    return tree


# --------------------------------------------------------------------------- #
# Main processing                                                             #
# --------------------------------------------------------------------------- #
def main():
    if not Path(INPUT_PPTX).is_file():
        raise FileNotFoundError(f'"{INPUT_PPTX}" not found')

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(INPUT_PPTX, "r") as zin:
            zin.extractall(tmpdir)

        ppt_path = Path(tmpdir) / "ppt"
        slides_dir = ppt_path / "slides"
        rels_dir = slides_dir / "_rels"

        pres_xml_path = ppt_path / "presentation.xml"
        pres_tree = ET.parse(pres_xml_path)
        pres_root = pres_tree.getroot()
        sldIdLst = pres_root.find("p:sldIdLst", namespaces=NS)

        pres_rels_path = ppt_path / "_rels" / "presentation.xml.rels"
        pres_rels_tree = ET.parse(pres_rels_path)
        pres_rels_root = pres_rels_tree.getroot()

        relinfo = {rel.get("Id"): (rel.get("Target"), rel.get("Type")) for rel in pres_rels_root}

        max_slide_num = 0
        for tgt, typ in relinfo.values():
            if typ.endswith("/slide"):
                m = re.search(r"/slide(\d+)\.xml$", tgt)
                if m:
                    max_slide_num = max(max_slide_num, int(m.group(1)))

        max_sldId = max(int(el.get("id")) for el in sldIdLst)

        for sldId in list(sldIdLst):
            relId = sldId.get(f'{{{NS["r"]}}}id')
            tgt, typ = relinfo[relId]
            if not typ.endswith("/slide"):
                continue

            slide_path = slides_dir / Path(tgt).name
            slide_tree = ET.parse(slide_path)
            shapes = collect_shapes(slide_tree)
            events = collect_visibility_events(slide_tree)

            if not events:
                materialise_snapshot(slide_tree, {spid: True for spid in shapes}).write(
                    slide_path, encoding="utf-8", xml_declaration=True
                )
                continue

            snapshots = build_snapshots(shapes, events)

            # step0 で差し替え
            materialise_snapshot(slide_tree, snapshots[0]).write(slide_path, encoding="utf-8", xml_declaration=True)

            orig_rels_path = rels_dir / f"{slide_path.name}.rels"
            for visible in snapshots[1:]:
                max_slide_num += 1
                new_slide_name = f"slide{max_slide_num}.xml"
                new_slide_path = slides_dir / new_slide_name
                materialise_snapshot(slide_tree, visible).write(new_slide_path, encoding="utf-8", xml_declaration=True)

                if orig_rels_path.exists():
                    shutil.copy(orig_rels_path, rels_dir / f"{new_slide_name}.rels")

                # 関連付けと sldId を追加
                new_relId = next_numeric_id(relinfo.keys(), "rId")
                rel_el = ET.Element(
                    "Relationship",
                    Id=new_relId,
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                    Target=f"slides/{new_slide_name}",
                )
                pres_rels_root.append(rel_el)
                relinfo[new_relId] = (rel_el.get("Target"), rel_el.get("Type"))

                max_sldId += 1
                new_sldId_el = ET.Element(ET.QName(NS["p"], "sldId"), id=str(max_sldId))
                new_sldId_el.set(ET.QName(NS["r"], "id"), new_relId)
                sldId.addnext(new_sldId_el)
                sldId = new_sldId_el  # ポインタ更新

        pres_tree.write(pres_xml_path, encoding="utf-8", xml_declaration=True)
        pres_rels_tree.write(pres_rels_path, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(OUTPUT_PPTX, "w", zipfile.ZIP_DEFLATED) as zout:
            for root, _, files in os.walk(tmpdir):
                for filename in files:
                    abs_path = Path(root) / filename
                    rel_path = abs_path.relative_to(tmpdir)
                    zout.write(abs_path, rel_path.as_posix())

    print(f'✅ Expanded presentation saved to "{OUTPUT_PPTX}"')


if __name__ == "__main__":
    main()
