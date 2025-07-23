#!/usr/bin/env python3
"""
pptx_split_anim_inout.py
-------------------
* input.pptx を読み込み
* entrance / exit アニメ付きシェイプを検出
* スライドを「Before（入口前）」＋「After（出口後）」の 2 枚に静的分割
* 出力ファイルを out_split.pptx として保存

依存:
    pip install python-pptx lxml
"""

from copy import deepcopy
from pathlib import Path, PurePosixPath
from zipfile import ZipFile
import posixpath
from typing import Set, Tuple

from lxml import etree
from pptx import Presentation
from pptx.slide import Slide  # ← これを既存の import 群に追加

# ---------- 設定 ---------- #
INPUT = Path(__file__).with_name("input.pptx")
OUTPUT = Path(__file__).with_name("out_split.pptx")

PML = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


# ---------- ユーティリティ ---------- #
def _xml(zf: ZipFile, part: str) -> etree._Element:
    return etree.fromstring(zf.read(part))


def _slide_xml_parts(zf: ZipFile) -> list[str]:
    """ppt/slides/slideX.xml の ZIP 内パスをソートして返す"""
    return sorted(p for p in zf.namelist() if p.startswith("ppt/slides/slide") and p.endswith(".xml"))


# def _entrance_exit_ids(slide_xml: etree._Element) -> Tuple[Set[str], Set[str]]:
#     """
#     <p:animEffect presetClass="entr|exit"> を走査し，
#     spTgt/@spid (shape id) をセットで返す
#     """
#     entr, exit_ = set(), set()
#     for anim in slide_xml.xpath(".//p:animEffect", namespaces=PML):
#         cls = anim.get("presetClass")  # 'entr', 'exit', 'emph' など
#         spids = {tgt.get("spid") for tgt in anim.xpath(".//p:spTgt", namespaces=PML)}
#         if cls == "entr":
#             entr.update(spids)
#         elif cls == "exit":
#             exit_.update(spids)
#     return entr, exit_


def _entrance_exit_ids(slide_xml: etree._Element) -> Tuple[Set[str], Set[str]]:
    entr, exit_ = set(), set()

    # 1) animEffect ベース（既存）
    for anim in slide_xml.xpath(".//p:animEffect", namespaces=PML):
        cls = anim.get("presetClass")
        flt = anim.get("filter", "")
        tgt_ids = {t.get("spid") for t in anim.xpath(".//p:spTgt", namespaces=PML)}
        if cls == "entr" or flt.startswith("in:"):
            entr.update(tgt_ids)
        elif cls == "exit" or flt.startswith("out:"):
            exit_.update(tgt_ids)

    # 2) set (opacity → 0) ベース（既存）
    for s in slide_xml.xpath(".//p:set[@to='0']", namespaces=PML):
        exit_.update({t.get("spid") for t in s.xpath(".//p:spTgt", namespaces=PML)})

    # 3) clickEffect ベース
    for ctn in slide_xml.xpath(".//p:cTn[@nodeType='clickEffect']", namespaces=PML):
        # presetClass で入口か出口かを判別
        cls = ctn.get("presetClass")
        tgt_ids = {t.get("spid") for t in ctn.xpath(".//p:spTgt", namespaces=PML)}

        if cls == "entr":
            entr.update(tgt_ids)
        elif cls == "exit":
            exit_.update(tgt_ids)
        else:
            # presetClass が無いケース → visibility=hidden / visible で判定
            to_hidden = ctn.xpath(
                ".//p:set[p:attrName='style.visibility']/p:to/p:strVal[@val='hidden']", namespaces=PML
            )
            if to_hidden:
                exit_.update(tgt_ids)
            else:
                entr.update(tgt_ids)

    return entr, exit_


# def _clone_slide(prs: Presentation, src_slide) -> Slide:
#     """python-pptx に clone API が無いので XML 丸ごと deep copy"""
#     blank = prs.slide_layouts[6]  # 空白レイアウト
#     new_slide = prs.slides.add_slide(blank)
#     # deep copy 元 XML を置換
#     new_slide._element.getparent().replace(new_slide._element, deepcopy(src_slide._element))
#     return new_slide


def _clone_slide(prs: Presentation, src: Slide) -> Slide:
    """
    - 空白レイアウトで新スライドを作成
    - src.shapes の XML ノードを deep copy して挿入
    - ノートやタイムラインはコピーしない（必要なら追加実装）
    """
    blank = prs.slide_layouts[6]  # 6 = Title Only / Blank など空白
    dst = prs.slides.add_slide(blank)

    # 背景やマスター依存スタイルが要る場合はここでコピーする
    #   dst.background = deepcopy(src.background)  # ←背景色だけならこの行でOK
    #
    # Shapes を丸ごと複製
    for shape in src.shapes:
        new_el = deepcopy(shape.element)
        # extLst の直前 (= spTree の末尾) に挿入
        dst.shapes._spTree.insert_element_before(new_el, "p:extLst")

    return dst


def _drop_shapes(slide, ids: Set[str]) -> None:
    """shape.id が ids に含まれる要素を削除"""
    for shp in list(slide.shapes):  # list() でコピーしてからでないと iterator 崩れる
        if str(shp.element.get("id")) in ids:
            shp.element.getparent().remove(shp.element)


# ---------- メイン ---------- #
def split_animation_slides(in_path: Path, out_path: Path) -> None:
    prs = Presentation(in_path)  # python-pptx オブジェクト
    with ZipFile(in_path) as zf:  # XML 低レイヤ解析
        slide_parts = _slide_xml_parts(zf)

        # enumerate は 0-origin なので +1 が pptx の SlideIndex
        for idx, slide_part in enumerate(slide_parts, start=1):
            xml_root = _xml(zf, slide_part)
            entr_ids, exit_ids = _entrance_exit_ids(xml_root)

            # アニメなしならスキップ
            if not (entr_ids or exit_ids):
                continue

            src_slide = prs.slides[idx - 1]

            # --- Before（入口前）: entrance シェイプを除去 --- #
            before_slide = _clone_slide(prs, src_slide)
            _drop_shapes(before_slide, entr_ids)

            # --- After（出口後）: exit シェイプを除去 --- #
            after_slide = _clone_slide(prs, src_slide)
            _drop_shapes(after_slide, exit_ids)

            # (Optional) どちらにも ' (before)' / ' (after)' タイトルを付けたい場合
            # if before_slide.shapes.title:
            #     before_slide.shapes.title.text += " (before)"
            # if after_slide.shapes.title:
            #     after_slide.shapes.title.text += " (after)"

    prs.save(out_path)
    print(f"✅ 分割完了: {out_path.absolute()}")


# ---------- 実行部 ---------- #
if __name__ == "__main__":
    if not INPUT.exists():
        raise SystemExit(f"❌ {INPUT} not found")
    split_animation_slides(INPUT, OUTPUT)
