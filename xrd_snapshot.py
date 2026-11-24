# -*- coding: utf-8 -*-
"""把 Origin 里当前这张图的所有样式读出来（快照 / 对比）

用法：
    python xrd_snapshot.py                 # 打印 + 存 _snapshot.json
    python xrd_snapshot.py --save base.json
    python xrd_snapshot.py --diff base.json   # 和之前存的快照比，只列变化

为什么要这个：用户在 Origin 窗口里手调完，我读一次快照就知道他改了哪几处
（颜色、线宽、字号、轴范围、图例位置…），再把这些值回写进 xrd_plot.py。
"""
import argparse
import json
import os
import sys

import win32com.client as w

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
NUM_KEYS = [
    ("layer.x.from", "X 下限"), ("layer.x.to", "X 上限"),
    ("layer.x.inc", "X 主刻度"), ("layer.x.minorTicks", "X 次刻度数"),
    ("layer.y.from", "Y 下限"), ("layer.y.to", "Y 上限"),
    ("layer.x.showAxes", "X 轴显示"), ("layer.y.showAxes", "Y 轴显示"),
    ("layer.x.label.pt", "X 刻度字号"), ("layer.y.label.pt", "Y 刻度字号"),
    ("layer.x.label.font", "X 刻度字体号"), ("layer.y.label.font", "Y 刻度字体号"),
    ("xb.fsize", "X 标题字号"), ("yl.fsize", "Y 标题字号"),
    ("legend.fsize", "图例字号"), ("legend.showframe", "图例边框"),
]


def num(name):
    try:
        return round(float(app.LTVar(name)), 3)
    except Exception:
        return None


def take(graph=None):
    global app
    app = w.Dispatch("Origin.ApplicationSI")
    if graph:
        app.Execute('win -a "%s";' % graph)
    pg = app.ActivePage
    if pg.Type != 3:
        raise SystemExit("当前活动窗口不是图: %s" % pg.Name)
    pg.Layers(0).Activate()      # 多图层时 LTVar("layer.*") 读的是活动图层
    ly = pg.Layers(0)
    snap = {"graph": pg.Name,
            "page_in": [round(pg.Width, 3), round(pg.Height, 3)],
            "layer_pct": [round(ly.GetNumProp(k), 2)
                          for k in ("Left", "Top", "Width", "Height")],
            "props": {}, "plots": [], "texts": []}
    for k, label in NUM_KEYS:
        snap["props"][label] = num(k)
    for i in range(ly.DataPlots.Count):
        dp = ly.DataPlots(i)
        c = int(dp.GetNumProp("color"))
        snap["plots"].append({
            "dataset": dp.GetDatasetName(),
            "color": "#%02X%02X%02X" % (c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF),
            "width_pt": round(float(dp.GetNumProp("line.width")), 2),
        })
    for i in range(ly.GraphObjects.Count):
        o = ly.GraphObjects(i)
        if o.Name.startswith("__") or not getattr(o, "Text", ""):
            continue
        snap["texts"].append({
            "name": o.Name, "text": o.Text,
            "left": o.Left, "top": o.Top, "w": o.Width, "h": o.Height,
            "rotate": round(float(o.GetNumProp("rotate")), 1),
        })
    return snap


def show(snap):
    print("图: %s   页面 %.2f x %.2f 英寸" % (snap["graph"], *snap["page_in"]))
    print("图层(%%): L=%.2f T=%.2f W=%.2f H=%.2f" % tuple(snap["layer_pct"]))
    for k, v in snap["props"].items():
        print("  %-12s %s" % (k, v))
    print("曲线:")
    for p in snap["plots"]:
        print("  %-14s %s  %.2f pt" % (p["dataset"], p["color"], p["width_pt"]))
    print("文字对象:")
    for t in snap["texts"]:
        print("  %-10s %-28s L=%s T=%s W=%s H=%s rot=%s"
              % (t["name"], t["text"].replace("\r\n", " / ")[:28],
                 t["left"], t["top"], t["w"], t["h"], t["rotate"]))


def diff(old, new):
    changed = 0
    for k in new["props"]:
        a, b = old["props"].get(k), new["props"].get(k)
        if a != b:
            print("  %-12s %s -> %s" % (k, a, b))
            changed += 1
    if old["page_in"] != new["page_in"]:
        print("  页面尺寸      %s -> %s" % (old["page_in"], new["page_in"]))
        changed += 1
    if old["layer_pct"] != new["layer_pct"]:
        print("  图层几何      %s -> %s" % (old["layer_pct"], new["layer_pct"]))
        changed += 1
    for i, p in enumerate(new["plots"]):
        q = old["plots"][i] if i < len(old["plots"]) else {}
        for k in ("color", "width_pt"):
            if q.get(k) != p.get(k):
                print("  曲线%d %-8s %s -> %s" % (i + 1, k, q.get(k), p.get(k)))
                changed += 1
    oldt = {t["name"]: t for t in old["texts"]}
    for t in new["texts"]:
        o = oldt.get(t["name"])
        if o is None:
            print("  新增文字      %s = %r @ (%s,%s)" % (t["name"], t["text"], t["left"], t["top"]))
            changed += 1
            continue
        for k in ("text", "left", "top", "w", "h", "rotate"):
            if o.get(k) != t.get(k):
                print("  文字 %-8s %-6s %r -> %r" % (t["name"], k, o.get(k), t.get(k)))
                changed += 1
    print("共 %d 处变化" % changed)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--save", default="_snapshot.json")
    ap.add_argument("--diff", default=None)
    ap.add_argument("--graph", default=None)
    ap.add_argument("--png", default="_snapshot.png")
    a = ap.parse_args(argv)
    snap = take(a.graph)
    show(snap)
    with open(os.path.join(HERE, a.save), "w", encoding="utf-8") as f:
        json.dump(snap, f, ensure_ascii=False, indent=2)
    print("快照已存:", a.save)
    if a.diff:
        with open(os.path.join(HERE, a.diff), encoding="utf-8") as f:
            old = json.load(f)
        print("=== 和 %s 比 ===" % a.diff)
        diff(old, snap)
    # 顺手导一张预览，方便肉眼看
    if a.png:
        p = os.path.join(HERE, a.png)
        if os.path.exists(p):
            os.remove(p)
        app.Execute('win -a "%s";' % snap["graph"])
        pw, ph = app.ActivePage.Width, app.ActivePage.Height
        px = 1200
        app.Execute('expGraph type:=png path:="%s" filename:="%s" overwrite:=replace '
                    'tr1.unit:=2 tr1.Width:=%d tr1.Height:=%d;'
                    % (HERE.replace("\\", "/"), os.path.splitext(a.png)[0],
                       px, int(round(px * ph / pw))))
        print("预览:", a.png)
        # 渲染里实际出现的彩色（防止 color 属性被主题覆盖时读不到真色）
        try:
            import numpy as np
            from PIL import Image
            arr = np.array(Image.open(p).convert("RGB")).reshape(-1, 3).astype(int)
            sat = arr.max(1) - arr.min(1)
            col = arr[(sat > 25) & (arr.max(1) > 60)]
            uniq, cnt = np.unique(col, axis=0, return_counts=True)
            top = [{"hex": "#%02X%02X%02X" % tuple(uniq[i]), "px": int(cnt[i])}
                   for i in np.argsort(-cnt)[:6]]
            snap["rendered_colors"] = top
            print("渲染实际用色:", ", ".join("%s(%d)" % (t["hex"], t["px"]) for t in top))
            with open(os.path.join(HERE, a.save), "w", encoding="utf-8") as f:
                json.dump(snap, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("取色失败:", e)
    return 0


if __name__ == "__main__":
    sys.exit(main())
