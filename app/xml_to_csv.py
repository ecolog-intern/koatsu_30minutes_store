# -*- coding: utf-8 -*-
"""
XML(中部/W4/0120のSBD-MSG形式) -> CSV 変換
- 名前空間有無に頑強（ローカルタグ名で探索）
- 1ファイル内の JPMR00010 グループごとに JP06219 を拾い、
  その配下の JPMR00011 明細を1行ずつ出力
- 出力先: out_dirを指定したらそこに<xml名>.csv、NoneならXMLと同じ場所に出力
- 戻り値: 生成した CSV のフルパス（str）
"""

import csv
from pathlib import Path
from xml.etree import ElementTree as ET

def _local(tag: str) -> str:
    return tag.split("}", 1)[-1] if tag and "}" in tag else (tag or "")

def _txt_any(parent, child_localname: str, default: str = "") -> str:
    """親以下から child_localname に一致する最初の子要素テキストを取得（名前空間無視）"""
    if parent is None:
        return default
    for e in parent:
        if _local(e.tag) == child_localname:
            return (e.text or "").strip()
    e = next((n for n in parent.iter() if _local(n.tag) == child_localname), None)
    return (e.text or "").strip() if (e is not None and e.text is not None) else default

def xml_to_csv(xml_path: str, out_dir: str | None = None) -> str:
    xml_p = Path(xml_path)
    if not xml_p.exists():
        raise FileNotFoundError(f"XMLが見つかりません: {xml_p}")

    # 出力先決定
    if out_dir:
        out_dir_p = Path(out_dir)
        out_dir_p.mkdir(parents=True, exist_ok=True)
        out_csv = out_dir_p / (xml_p.stem + ".csv")
    else:
        out_csv = xml_p.with_suffix(".csv")

    # 解析（未完成XMLならここで例外）
    tree = ET.parse(xml_p)
    root = tree.getroot()

    # 共通部（最初に見つかったもの）
    jpmgrp  = next((n for n in root.iter() if _local(n.tag)=="JPMGRP"), None)
    jpmgh   = next((n for n in (jpmgrp or root).iter() if _local(n.tag)=="JPMGH"), None)
    jptrm   = next((n for n in (jpmgrp or root).iter() if _local(n.tag)=="JPTRM"), None)

    common_jpmgh = {k: _txt_any(jpmgh, k) for k in
                    ["JPC03","JPC06","JPC09","JPC10","JPC11","JPC12","JPC14","JPC19","JPC21"]}
    common_jptrm = {k: _txt_any(jptrm, k) for k in
                    ["JP00002","JP06110","JP06111","JP06112","JP06113","JP06114","JP06115","JP06116"]}

    # JPMR00010 グループ（= JP06219が載ってる）
    groups = [n for n in (jpmgrp or root).iter() if _local(n.tag)=="JPMR00010"]

    header = [
        # JPMGH
        "JPC03","JPC06","JPC09","JPC10","JPC11","JPC12","JPC14","JPC19","JPC21",
        # JPTRM
        "JP00002","JP06110","JP06111","JP06112","JP06113","JP06114","JP06115","JP06116",
        # JPMR00010
        "JP06219",
        # JPMR00011（明細）
        "JP06400","JP06119","JP06120","JP06121","JP06122","JP06123"
    ]

    rows = []
    for g in groups:
        jp06219 = _txt_any(g, "JP06219")
        jpm00011 = next((n for n in g.iter() if _local(n.tag)=="JPM00011"), None)
        if jpm00011 is None:
            continue
        details = [n for n in jpm00011.iter() if _local(n.tag)=="JPMR00011"]
        for rec in details:
            row = []
            row.extend([common_jpmgh[k] for k in ["JPC03","JPC06","JPC09","JPC10","JPC11","JPC12","JPC14","JPC19","JPC21"]])
            row.extend([common_jptrm[k] for k in ["JP00002","JP06110","JP06111","JP06112","JP06113","JP06114","JP06115","JP06116"]])
            row.append(jp06219)
            row.extend([
                str(_txt_any(rec,"JP06400")),
                _txt_any(rec,"JP06119"),
                _txt_any(rec,"JP06120"),
                _txt_any(rec,"JP06121"),
                _txt_any(rec,"JP06122"),
                _txt_any(rec,"JP06123"),
            ])
            rows.append(row)

    # CSV書き込み（Excel向けにUTF-8 BOM）
    with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)

    return str(out_csv)
