# -*- coding: utf-8 -*-
"""JCPDS 卡片库维护：从 JADE 导出的 PDF#xxxx.txt 导入，写进 xrd_cards.json

用法：
    python xrd_cards.py --list
    python xrd_cards.py --import "/path/to/PDF#05-0667.txt" [--key Cu2O]
    python xrd_cards.py --show Cu2O
    python xrd_cards.py --search "cuprite"          # 在 COD 里按名字找结构
    python xrd_cards.py --from-cod 1000063 --key Cu2O   # 下 CIF 算谱导入

两种来源：
- JADE 导出的 PDF#xxxx.txt = PDF 卡片原值，可引用，**出图用这个**
- COD 的 CIF 算出来的谱 = 2θ 准、强度是理论值（和卡片能差 10-30%），**只用来定物相**

为什么要有这个：手抄卡片强度会抄错（Cu2O 就把 (110)=9 抄成了 100），
JADE 导出的 txt 里 2θ / d / I(f) / hkl 都是卡片原值，直接解析最稳。

JADE 导出格式（每张卡一个 txt）：
    PDF#05-0667: QM=Common(+); d=Diffractometer; I=(Unknown)
    Cuprite, syn
    Cu2+1O
    Radiation=CuKa1  Lambda=1.5406
    ...
    Cubic,  Pn-3m(224)  Z=2
    CELL: 4.2696 x 4.2696 x 4.2696 <90.0 x 90.0 x 90.0>
    ...
    2-Theta    d(A)   I(f)  ( h k l)   Theta  1/(2d)   2pi/d  n^2
     29.554  3.0200    9.0  ( 1 1 0)  14.777  0.1656  2.0805    2
"""
import argparse
import json
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
CARDS_JSON = os.path.join(HERE, "xrd_cards.json")

RE_CARD = re.compile(r"PDF#\s*([0-9]{2}-[0-9]{4})")
RE_PEAK = re.compile(r"^\s*(\d+\.\d+)\s+([\d.]+)\s+([\d.]+)\s+\(\s*([\d\s]+?)\s*\)")
RE_CELL = re.compile(r"CELL:\s*([\d.]+)\s*x\s*([\d.]+)\s*x\s*([\d.]+)")


def clean_formula(s):
    """Cu2+1O -> Cu2O；Co+2O-2 -> CoO。去掉 JADE 的价态标记。"""
    s = s.strip().split("\t")[0].split()[0] if s.strip() else ""
    return re.sub(r"[+-]\d+", "", s)


def parse_jade(path):
    """解析 JADE 导出的 PDF 卡片 txt，返回卡片 dict。"""
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        lines = [ln.rstrip("\n") for ln in f]

    card = {"card": None, "name": "", "formula": "", "system": "",
            "spacegroup": "", "lattice": {}, "peaks": [],
            "intensity_source": "card (JADE 导出)", "source": "jade",
            "source_file": path}
    head_idx = None
    for i, ln in enumerate(lines):
        m = RE_CARD.search(ln)
        if m:
            card["card"] = "PDF#" + m.group(1)
            head_idx = i
            break
    if head_idx is None:
        raise SystemExit("没找到 PDF# 卡片号: %s" % path)

    rest = lines[head_idx + 1:]
    if rest:
        card["name"] = rest[0].strip()
    if len(rest) > 1:
        card["formula"] = clean_formula(rest[1])

    for ln in rest:
        m = RE_CELL.search(ln)
        if m:
            a, b, c = [float(v) for v in m.groups()]
            card["lattice"] = {"a": a, "b": b, "c": c}
        low = ln.strip().lower()
        if low.startswith(("cubic", "tetragonal", "hexagonal", "orthorhombic",
                           "monoclinic", "triclinic", "trigonal", "rhombohedral")):
            parts = [p.strip() for p in ln.split(",")]
            card["system"] = parts[0].strip().lower()
            if len(parts) > 1:
                card["spacegroup"] = parts[1].split("\t")[0].strip()

    for ln in rest:
        m = RE_PEAK.match(ln)
        if not m:
            continue
        t2, d, inten, hkl = m.groups()
        card["peaks"].append([float(t2), float(inten),
                              re.sub(r"\s+", "", hkl)])
    if not card["peaks"]:
        raise SystemExit("没解析到峰: %s" % path)
    return card


def cod_search(text, limit=10):
    """在 COD 里按矿物名/化学式搜结构，返回 [(id, title), ...]。"""
    import urllib.parse
    import urllib.request
    url = ("https://www.crystallography.net/cod/result?text=%s&format=json"
           % urllib.parse.quote(text))
    with urllib.request.urlopen(url, timeout=30) as r:
        rows = json.loads(r.read().decode("utf-8"))
    out = []
    for row in rows[:limit]:
        out.append((str(row.get("file")), (row.get("title") or "").strip()))
    return out


def cod_fetch(cod_id, out_dir=None):
    """下载 COD 的 CIF，返回本地路径。"""
    import urllib.request
    out_dir = out_dir or os.path.join(HERE, "cif")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "%s.cif" % cod_id)
    if not os.path.exists(path):
        url = "https://www.crystallography.net/cod/%s.cif" % cod_id
        urllib.request.urlretrieve(url, path)
    return path


def compute_pattern(cif_path, wavelength="CuKa1", two_theta=(10.0, 90.0)):
    """用 pymatgen 从 CIF 算粉末衍射谱，返回 (peaks, structure)。"""
    try:
        from pymatgen.analysis.diffraction.xrd import XRDCalculator
        from pymatgen.core import Structure
    except ImportError:
        raise SystemExit("没装 pymatgen：pip install pymatgen（算 CIF 谱要用）")
    s = Structure.from_file(cif_path)
    pat = XRDCalculator(wavelength=wavelength).get_pattern(
        s, two_theta_range=two_theta, scaled=True)
    peaks = []
    for t, inten, hkls in zip(pat.x, pat.y, pat.hkls):
        hkl = hkls[0]["hkl"] if isinstance(hkls, list) else hkls
        peaks.append([round(float(t), 3), round(float(inten), 1),
                      "".join(str(int(v)) for v in hkl) if isinstance(hkl, (tuple, list)) else str(hkl)])
    return peaks, s


def load_db():
    with open(CARDS_JSON, encoding="utf-8") as f:
        return json.load(f)


def save_db(db):
    with open(CARDS_JSON, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)
        f.write("\n")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--import", dest="imp", metavar="PDF_TXT")
    ap.add_argument("--key", default=None, help="卡片库里的键名，默认用化学式")
    ap.add_argument("--show", default=None)
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--search", default=None, help="在 COD 里按名字搜结构")
    ap.add_argument("--from-cod", dest="from_cod", default=None, metavar="COD_ID")
    a = ap.parse_args(argv)

    db = load_db()
    if a.list:
        for k, v in db.items():
            if k.startswith("_"):
                continue
            print("%-8s %-12s %-28s %d 峰" % (k, v.get("card", ""),
                                             v.get("name", ""), len(v["peaks"])))
        return 0

    if a.show:
        v = db[a.show]
        print(json.dumps(v, ensure_ascii=False, indent=2))
        return 0

    if a.search:
        for cid, title in cod_search(a.search):
            print("%-10s %s" % (cid, title))
        return 0

    if a.from_cod:
        cif = cod_fetch(a.from_cod)
        peaks, s = compute_pattern(cif)
        key = a.key
        if not key:
            key = re.sub(r"[^A-Za-z0-9]", "", s.composition.reduced_formula)
        card = {
            "formula": s.composition.reduced_formula,
            "name": "computed from COD %s" % a.from_cod,
            "card": "COD#%s" % a.from_cod,
            "system": s.lattice._symmetry if hasattr(s.lattice, "_symmetry") else "",
            "spacegroup": str(s.get_space_group_info()[0]) if s.get_space_group_info() else "",
            "lattice": {"a": round(s.lattice.a, 4), "b": round(s.lattice.b, 4),
                        "c": round(s.lattice.c, 4)},
            "source": "cod",
            "source_file": cif,
            "intensity_source": "computed (pymatgen, CuKa1)",
            "peaks": [p for p in peaks if p[1] >= 0.5],
        }
        old = db.get(key, {})
        if old.get("label"):
            card["label"] = old["label"]
        db[key] = card
        save_db(db)
        print("已导入 COD#%s -> 键名 %s（计算谱，只能用来比对，不要当卡片引用）"
              % (a.from_cod, key))
        print("%-4s %8s %8s %6s" % ("#", "2θ(°)", "I", "hkl"))
        for i, (t2, inten, hkl) in enumerate(card["peaks"], 1):
            print("%-4d %8.3f %8.1f %6s" % (i, t2, inten, hkl))
        return 0

    if not a.imp:
        print(__doc__)
        return 1

    card = parse_jade(os.path.abspath(a.imp))
    key = a.key or card["formula"]
    if not key:
        raise SystemExit("化学式为空，请用 --key 指定键名")
    old = db.get(key, {})
    if old.get("label"):
        card["label"] = old["label"]
    elif re.search(r"2O$", key):
        card["label"] = key.replace("2O", r"\-2O")
    db[key] = card
    save_db(db)

    print("已导入 %s -> 键名 %s" % (card["card"], key))
    print("%-4s %8s %8s %6s" % ("#", "2θ(°)", "I(f)", "hkl"))
    for i, (t2, inten, hkl) in enumerate(card["peaks"], 1):
        print("%-4d %8.3f %8.1f %6s" % (i, t2, inten, hkl))
    print("卡片库:", CARDS_JSON)
    return 0


if __name__ == "__main__":
    sys.exit(main())
