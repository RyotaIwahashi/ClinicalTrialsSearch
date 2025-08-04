#!/usr/bin/env python3
"""
pptx_split_animation.py
~~~~~~~~~~~~~~~~~~~~~~~

End‑to‑end utility: read *input.pptx*, expand every animated slide
(click‑triggered visibility toggles) into a sequence of static slides,
and save as *output.pptx*.

Highlights
----------
* Keeps theme/master/notes/handout parts untouched.
* Only considers `<p:set>` animations whose attrName is *style.visibility*.
  Motion paths・emphasis etc. are ignored.
* Original animated slide is replaced by its first static snapshot
  (step0). Additional snapshots are appended immediately after it, so
  the logical order is maintained.
* Does **not** delete unused relationships (images that became hidden
  forever). PowerPoint tolerates dangling rels.

Dependencies
------------
`pip install lxml`

Tested with Python 3.9 / Office 2016.

Limitations
-----------
* Group shapes are treated as a unit (child visibility changes are not
  exploded individually).
* Does not touch slide layouts. If a placeholder gains visibility later,
  it is copied as a concrete shape in each snapshot.
* Re‑numbering is naive but works for default «slideX.xml» pattern.

Author: ChatGPT (o3 model)  — 2025‑07‑25
License: MIT
"""

import zipfile
import shutil
import os
import tempfile
import copy
import re
from pathlib import Path
from lxml import etree as ET

INPUT_PPTX = "input.pptx"
OUTPUT_PPTX = "output.pptx"

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
    """
    既存のIDリストから最大の数値部分を探し、1増やした新しいID文字列を返す。
    例: ["rId1", "rId2"] → "rId3"
    """
    nums = [int(re.sub(r"\D", "", s)) for s in existing if re.sub(r"\D", "", s)]
    return f"{prefix}{max(nums, default=0) + 1}"


def collect_shapes(slide_tree):
    """
    スライドXMLから全ての図形（shape）のspidと要素を辞書として収集する。

    args:
        slide_tree (lxml.etree._ElementTree): スライドのXMLツリー
    returns:
        dict: spidをキー、図形要素を値とする辞書
    例: {1: <p:sp>...</p:sp>, 2: <p:pic>...</p:pic>, ...}
    """
    shapes = {}
    for el in slide_tree.xpath("//p:sp | //p:pic | //p:graphicFrame | //p:grpSp", namespaces=NS):
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
    スライドXMLから「表示/非表示」アニメーションイベント(spid, visible)のリストを抽出する。

    args:
        slide_tree (lxml.etree._ElementTree): スライドのXMLツリー
    returns:
        list: (spid, visible) のタプルのリスト
        visible はアニメーション後の状態を示す。True なら表示、False なら非表示を意味する。
    例: [(1, True), (2, False), (3, True), ...]
    """
    events = []
    for set_el in slide_tree.xpath("//p:set", namespaces=NS):
        attr_names = set_el.xpath(".//p:attrNameLst/p:attrName/text()", namespaces=NS)

        # style.visibility を変える <p:set> が対象
        if "style.visibility" not in attr_names:
            continue

        spid_attr = set_el.xpath(".//p:spTgt/@spid", namespaces=NS)  # 図形を識別するID
        to_val = set_el.xpath(
            "./p:to/p:strVal/@val", namespaces=NS
        )  # アニメーション後の状態。表示:visible、非表示:hidden。
        if not spid_attr or not to_val:
            continue

        try:
            spid = int(spid_attr[0])
        except ValueError:
            continue

        # to_val[0] が "hidden" なら visible=False、それ以外は visible=True
        visible = to_val[0] != "hidden"
        events.append((spid, visible))
    return events


def build_snapshots(shapes, events):
    """
    図形の初期状態とイベント列から、各ステップごとの可視状態スナップショットのリストを作成する。

    args:
        shapes (dict): spidをキーとする図形の辞書
        events (list): (spid, visible) のタプルのリスト
    returns:
        list: 各ステップの可視状態を表す辞書のリスト
        各辞書は spid をキー、可視状態 (True/False) を値とする。
    例: [{1: True, 2: False}, {1: True, 2: True}, ...]
    """
    state = {spid: True for spid in shapes}
    first_seen = {}

    # 初期の可視状態を設定する
    for spid, visible in events:
        if spid not in first_seen:
            first_seen[spid] = visible
            if visible:
                # 後から表示されるため、初期状態では非表示にする
                state[spid] = False
            else:
                state[spid] = True
    snaps = [copy.deepcopy(state)]

    # イベントに基づいて各shapeの状態を更新
    for spid, visible in events:
        if spid in state:
            state[spid] = visible
        snaps.append(copy.deepcopy(state))
    return snaps


def materialise_snapshot(orig_tree, visible_map):
    """
    指定した可視状態(visible_map)に従い、スライドXMLから非表示図形を除去した新しいツリーを返す。
    アニメーション情報も削除する。
    """
    tree = copy.deepcopy(orig_tree)
    # Remove timing for static slide
    timing = tree.find(".//p:timing", namespaces=NS)
    if timing is not None:
        timing.getparent().remove(timing)

    for el in tree.xpath("//p:sp | //p:pic | //p:graphicFrame | //p:grpSp", namespaces=NS):
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
    """
    PPTXファイルを展開し、各スライドのアニメーションを静的スライドに分割して新しいPPTXを作成するメイン処理。
    """
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

        # プレゼンテーションのリレーションシップ(スライド, 画像などのパス)を収集
        relinfo = {rel.get("Id"): (rel.get("Target"), rel.get("Type")) for rel in pres_rels_root}

        # リレーションに新しいスライドを追加する用の最大スライド番号を取得(連番)
        max_slide_num = 0
        for tgt, typ in relinfo.values():
            if typ.endswith("/slide"):
                m = re.search(r"/slide(\d+)\.xml$", tgt)
                if m:
                    num = int(m.group(1))
                    max_slide_num = max(max_slide_num, num)

        # スライドに関連付けられたIDの最大値を取得(連番)
        max_sldId = max(int(el.get("id")) for el in sldIdLst)

        for sldId in list(sldIdLst):
            relId = sldId.get(f'{{{NS["r"]}}}id')
            tgt, typ = relinfo[relId]
            if not typ.endswith("/slide"):
                continue  # スライドのみを処理対象とする

            slide_path = slides_dir / Path(tgt).name
            slide_tree = ET.parse(slide_path)
            shapes = collect_shapes(slide_tree)
            events = collect_visibility_events(slide_tree)

            # スライドにアニメーションがない場合は、可視状態スナップショットをそのまま保存
            if not events:
                static = materialise_snapshot(slide_tree, {spid: True for spid in shapes})
                static.write(slide_path, encoding="utf-8", xml_declaration=True)
                continue

            snapshots = build_snapshots(shapes, events)

            # 当該スライドを最初のスナップショットで静的スライドに置き換える
            materialise_snapshot(slide_tree, snapshots[0]).write(slide_path, encoding="utf-8", xml_declaration=True)

            # TODO:次はここから。どのようにしてページ番号を設定し、スライドの後ろに追加しているのかを整理する。
            orig_rels_path = rels_dir / f"{slide_path.name}.rels"
            for visible in snapshots[1:]:
                max_slide_num += 1
                new_slide_name = f"slide{max_slide_num}.xml"
                new_slide_path = slides_dir / new_slide_name

                materialise_snapshot(slide_tree, visible).write(new_slide_path, encoding="utf-8", xml_declaration=True)

                if orig_rels_path.exists():
                    shutil.copy(orig_rels_path, rels_dir / f"{new_slide_name}.rels")

                # add new relationship
                new_relId = next_numeric_id(relinfo.keys(), "rId")
                rel_el = ET.Element(
                    "Relationship",
                    Id=new_relId,
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                    Target=f"slides/{new_slide_name}",
                )
                pres_rels_root.append(rel_el)
                relinfo[new_relId] = (rel_el.get("Target"), rel_el.get("Type"))

                # add new sldId after current position
                max_sldId += 1
                new_sldId_el = ET.Element(ET.QName(NS["p"], "sldId"), id=str(max_sldId))
                new_sldId_el.set(ET.QName(NS["r"], "id"), new_relId)
                sldId.addnext(new_sldId_el)
                sldId = new_sldId_el  # advance reference

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
