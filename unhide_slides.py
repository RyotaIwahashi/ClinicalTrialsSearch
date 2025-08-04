#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from io import BytesIO
from lxml import etree

INPUT_PPTX = "hide_input.pptx"
OUTPUT_PPTX = "hide_output.pptx"

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _unhide_slide_xml(xml_bytes: bytes) -> tuple[bytes, bool]:
    """ppt/slides/slideX.xml を受け取り、show 属性を外す/true にして返す"""
    root = etree.fromstring(xml_bytes)
    changed = False

    # p:show と無名属性 show の両方を見る（生成ツールによって揺れるため）
    for attr_name in ("show", "{%s}show" % NS["p"]):
        val = root.get(attr_name)
        if val is not None and val.lower() in ("0", "false"):
            root.set(attr_name, "1")  # もしくは root.attrib.pop(attr_name)
            changed = True

    if changed:
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"), True
    return xml_bytes, False


def _unhide_presentation_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    """ppt/presentation.xml の <p:sldId hidden="1"> を解除"""
    root = etree.fromstring(xml_bytes)
    cnt = 0
    for sldId in root.xpath(".//p:sldId[@hidden='1']", namespaces=NS):
        del sldId.attrib["hidden"]
        cnt += 1
    if cnt:
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"), cnt
    return xml_bytes, 0


def unhide_all(input_pptx: Path, output_pptx: Path) -> tuple[int, int]:
    """
    すべての非表示スライドを表示に戻す。
    Returns: (slides_fixed, ids_fixed)
      slides_fixed ... slideX.xml 側で修正した枚数
      ids_fixed     ... presentation.xml 側で hidden を外した件数
    """
    slides_fixed = 0
    ids_fixed = 0

    with ZipFile(input_pptx, "r") as zin, ZipFile(output_pptx, "w", ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)

            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                data, changed = _unhide_slide_xml(data)
                if changed:
                    slides_fixed += 1

            elif name == "ppt/presentation.xml":
                data, fixed = _unhide_presentation_xml(data)
                ids_fixed += fixed

            # そのまま/書き換えた data を書き出す
            zout.writestr(name, data)

    return slides_fixed, ids_fixed


if __name__ == "__main__":
    s_fixed, id_fixed = unhide_all(INPUT_PPTX, OUTPUT_PPTX)
    print(f"Done: slides(show attr) fixed={s_fixed}, sldId(hidden) fixed={id_fixed}")
