# -*- coding: utf-8 -*-
"""XRD 堆叠图（Origin COM 路线，两层结构）：上层样品谱，下层标准卡片棒图

用法：
    python xrd_plot.py <data.txt> --cards Cu2O Cu [--peaks 3] [--smooth 15]

结构（2026-11-24 用户在 Origin 里手调后确认）：
- 图层1 样品：归一化 0-100 放在 28-128，Y 轴 0-140；方框轴、Y 无刻度无数字
- 图层2 卡片：与图层1 同矩形叠加，Y 轴 0-CARD_Y_MAX(1000) → 棒高 100 = 框高 10%，
              卡片强度直接用 PDF 卡片的相对强度 I/I0，X/Y 轴都关
- 图例只留样品名，右上（右边距 177、顶边距 318 设备单位）
- Y 轴标题是普通文字对象旋转 90°，框左侧 98 设备单位
- 页面裁到墨迹 + 四周 150 设备单位

配色：样品 #8E9CFF 2.0 pt；卡片 #FDC1C1 / #CEF297 3.0 pt

LabTalk 坑：`layerN.xxx` 前缀不可靠（会落到当前活动图层上），
所以一律先 `page.active = N` 再用 `layer.xxx`。

交付物：<名>.tif (300 dpi) + <名>.opju
"""
import argparse
import json
import os
import sys
import time

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from origin_plot import OriginPlot  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

KEEP_PAGE = "Book1"
MARGIN = 150
DPI = 300
CARDS_JSON = os.path.join(HERE, "xrd_cards.json")

PAGE_W, PAGE_H = 8.43, 7.40      # 英寸（用户手调定稿的页面）
LAYER = (8.39, 3.49, 86.73, 79.62)   # 图层矩形：左/上/宽/高（%）
CARD_Y_MAX = 1000.0
PAD_BOTTOM = 28.0
PAD_TOP = 12.0
LEGEND_PAD_R = 177
LEGEND_PAD_T = 318
YTITLE_GAP = 98
SAMPLE_COLOR = "142,156,255"
CARD_COLORS = ["253,193,193", "206,242,151", "180,180,180", "255,200,120"]
SAMPLE_W = 1000
STICK_W = 1500


def load_xy(path):
    xs, ys = [], []
    in_data = False
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            s = line.strip()
            if not s:
                continue
            if s.startswith("["):
                in_data = s.lower().startswith("[data")
                continue
            if s.startswith(";") or s.startswith("#"):
                continue
            parts = s.replace(",", " ").split()
            if len(parts) < 2:
                continue
            try:
                x, y = float(parts[0]), float(parts[1])
            except ValueError:
                continue
            if not in_data and not (s[0].isdigit() or s[0] in "+-."):
                continue
            xs.append(x)
            ys.append(y)
    if not xs:
        raise SystemExit("没有读到数据: %s" % path)
    order = np.argsort(xs)
    return np.array(xs)[order], np.array(ys, dtype=float)[order]


def load_card(name):
    with open(CARDS_JSON, encoding="utf-8") as f:
        db = json.load(f)
    if name not in db:
        raise SystemExit("卡片库里没有 %s，可选: %s"
                         % (name, ", ".join(k for k in db if not k.startswith("_"))))
    return db[name]


def build_sample(x, y, xlim):
    m = (x >= xlim[0]) & (x <= xlim[1])
    xs, ys = x[m], y[m]
    ymin, ymax = float(ys.min()), float(ys.max())
    ynorm = (ys - ymin) / max(ymax - ymin, 1e-9) * 100.0 + PAD_BOTTOM
    rows = [[float(a), float(b)] for a, b in zip(xs, ynorm)]
    return rows, (0.0, PAD_BOTTOM + 100.0 + PAD_TOP)


def build_cards(cards, xlim, base=-2.0):
    """每根棒写成「基线→顶→回基线」三个点。

    不能用 NaN 断线：Origin 在追加到图层的曲线上会「跨缺失值连线」，
    棒子会被斜线连成一串（实测）。改成回到基线后，连接段落在 y=base
    （在 y 轴下限之外），被图层裁掉，视觉上就是一根根独立竖棒。
    """
    rows = []
    for k, card in enumerate(cards):
        for t2, rel, _hkl in card["peaks"]:
            if not (xlim[0] <= t2 <= xlim[1]):
                continue
            for v in (base, float(rel), base):
                row = [float(t2)] + [float("nan")] * len(cards)
                row[1 + k] = v
                rows.append(row)
    rows.sort(key=lambda r: r[0])
    return rows


# ---------- 标注 ----------
def find_top_peaks(rows, n=3, min_prom=0.08):
    """从画图用的 (x, y) 里找最强的 n 个峰。"""
    import numpy as np
    from scipy.signal import find_peaks
    xs = np.array([r[0] for r in rows])
    ys = np.array([r[1] for r in rows])
    rng = float(ys.max() - ys.min()) or 1.0
    idx, _props = find_peaks(ys, prominence=min_prom * rng)
    order = np.argsort(ys[idx])[::-1][:n]
    return [(float(xs[idx[i]]), float(ys[idx[i]])) for i in order]


def match_card(two_theta, cards, tol=0.6):
    """给一个样品峰位，找最近的卡片线，返回 (物相文字, hkl, 卡片序号)。"""
    best = None
    for k, card in enumerate(cards):
        for t2, _rel, hkl in card["peaks"]:
            d = abs(t2 - two_theta)
            if d <= tol and (best is None or d < best[0]):
                best = (d, k, hkl)
    if best is None:
        return None
    _d, k, hkl = best
    return cards[k].get("label", cards[k]["formula"]), hkl, k


def add_text(op, name, text, size=22, font="Arial", color=None):
    """新建文字对象；Left/Top 用 COM 设（单位 1/600 英寸）。"""
    op.app.Execute('page.active = 1;')
    op.app.Execute('label -s -fs %g -fn %s -n %s "%s";' % (size, font, name, text))
    ly = op.app.ActivePage.Layers(0)
    for i in range(ly.GraphObjects.Count):
        o = ly.GraphObjects(i)
        if o.Name == name:
            if color is not None:
                o.SetNumProp("color", float(color))
            return o
    return None


def add_peak_labels(op, rows, cards, colors, n=3, tol=0.6):
    """标出最强 n 个峰的晶面（Cu (111) 之类），颜色跟卡片一致。"""
    L, T, W, H = layer_rect(op, 0)
    op.app.ActivePage.Layers(0).Activate()
    x0, x1 = op.app.LTVar('layer.x.from'), op.app.LTVar('layer.x.to')
    y0, y1 = op.app.LTVar('layer.y.from'), op.app.LTVar('layer.y.to')

    def px(v):
        return L + (v - x0) / (x1 - x0) * W

    def py(v):
        return T + (y1 - v) / (y1 - y0) * H

    done = 0
    for i, (x, y) in enumerate(find_top_peaks(rows, n + 3), start=1):
        if done >= n:
            break
        m = match_card(x, cards, tol)
        if m is None:
            continue
        done += 1
        lab, hkl, k = m
        obj = add_text(op, "pk%d" % done, "%s (%s)" % (lab, hkl), color=_rgb(colors[k]))
        if obj is None:
            continue
        obj.Left = int(px(x) - obj.Width / 2.0)
        obj.Top = int(py(y) - obj.Height - 30)


def add_card_labels(op, cards, colors, rows=None, xlim=None, card_y_max=CARD_Y_MAX,
                    pad_r=LEGEND_PAD_R, size=20, gap=2):
    """卡片号写在「样品曲线下方、PDF 棒图上方」那条空白带的右侧，右对齐。

    带子上边界取**右侧那一段**曲线的最低点（不是全谱最低点）—— 右边曲线本来就低，
    按全谱最低点摆会顶到线上。字号默认 20（用户要求不小于 20）。
    """
    L, T, W, H = layer_rect(op, 0)
    op.app.ActivePage.Layers(0).Activate()
    y0 = op.app.LTVar('layer.y.from')
    y1 = op.app.LTVar('layer.y.to')

    def py(v):
        return T + (y1 - v) / (y1 - y0) * H

    top_data = PAD_BOTTOM - 1.0          # 兜底
    if rows and xlim:
        x_cut = xlim[0] + 0.72 * (xlim[1] - xlim[0])
        tail = [r[1] for r in rows if r[0] >= x_cut]
        if tail:
            top_data = min(tail) - 1.0
    top1 = py(top_data)
    for k, card in enumerate(cards):
        lab = card.get("label", card["formula"])
        obj = add_text(op, "cardlab%d" % (k + 1),
                       "%s %s" % (lab, card["card"]), size=size,
                       color=_rgb(colors[k]))
        if obj is None:
            continue
        obj.Left = int(L + W - pad_r - obj.Width)
        obj.Top = int(top1 + k * (obj.Height + gap))
    return True


def place_ytitle_gap(op, gap=98):
    """把内置 Y 轴标题挪到「右边缘距轴线 gap」设备单位。

    坑：yl.x 的零点位置跟图的历史有关（同一台机器上，有的图 yl.x≈0 就贴在轴上，
    有的要 +17 才进左边距，有的要 -2）。所以不能写死窗口，用「朝目标方向逐步逼近 +
    跨过目标后插值」：
      - 先量 yl.x=0 的位置，判断该往正还是负走
      - 每步走 1，跨过目标就插值；走出左边距（bbox=None）就步长减半退回
    """
    op.app.ActivePage.Layers(0).Activate()
    L, T, W, H = layer_rect(op, 0)
    target = L - gap

    def right_at(x):
        op.app.Execute('yl.x = %g;' % x)
        tb = op._left_title_bbox()
        return None if tb is None else tb[2]

    x, r = 0.0, right_at(0.0)
    if r is None:                       # yl.x=0 时不在左边距里，先找回来
        for probe in (-2, 2, -5, 5, -10, 10, -16, 16):
            r = right_at(probe)
            if r is not None:
                x = probe
                break
        if r is None:
            return False
    if abs(r - target) <= 2:
        return True

    direction = -1.0 if target < r else 1.0
    step = 1.0
    prev_x, prev_r = x, r
    for _ in range(16):
        x = prev_x + direction * step
        r = right_at(x)
        if r is None:                   # 走出左边距了，退半步
            step /= 2.0
            if step < 0.05:
                op.app.Execute('yl.x = %g;' % prev_x)
                return True
            continue
        if (r - target) * direction >= 0:
            if abs(r - prev_r) > 1e-6:
                x = prev_x + (target - prev_r) * (x - prev_x) / (r - prev_r)
            op.app.Execute('yl.x = %g;' % x)
            return True
        prev_x, prev_r = x, r
    op.app.Execute('yl.x = %g;' % prev_x)
    return True


def layer_rect(op, k):
    pg = op.app.ActivePage
    PW, PH = pg.Width, pg.Height
    ly = pg.Layers(k)
    L, T, W, H = [ly.GetNumProp(x) for x in ("Left", "Top", "Width", "Height")]
    return (L / 100.0 * PW * 600, T / 100.0 * PH * 600,
            W / 100.0 * PW * 600, H / 100.0 * PH * 600)


def set_rect(op, k, L, T, W, H, pw, ph):
    op.app.Execute('page.active = %d;' % (k + 1))
    op.app.Execute('layer.left=%g; layer.top=%g; layer.width=%g; layer.height=%g;'
                   % (L / pw * 100, T / ph * 100, W / pw * 100, H / ph * 100))


def trim_layers(op, margin=MARGIN, iters=3):
    for _ in range(iters):
        ink = op._probe_ink()
        if ink is None:
            return op.page_size()
        pw, ph = op.page_size()
        L, T, W, H = layer_rect(op, 0)
        il, it, ir, ib = ink
        npw = (ir - il) + 2 * margin
        nph = (ib - it) + 2 * margin
        if (abs(npw - pw) < 2 and abs(nph - ph) < 2
                and abs(il - margin) < 2 and abs(it - margin) < 2):
            break
        op.app.Execute('page.width=%g; page.height=%g;' % (npw, nph))
        nL, nT = L + (margin - il), T + (margin - it)
        for k in range(op.app.ActivePage.Layers.Count):
            set_rect(op, k, nL, nT, W, H, npw, nph)
    return op.page_size()


def place_legend(op, pad_r=LEGEND_PAD_R, pad_t=LEGEND_PAD_T):
    pg = op.app.ActivePage
    ly = pg.Layers(0)
    lgn = None
    for i in range(ly.GraphObjects.Count):
        if ly.GraphObjects(i).Name == "Legend":
            lgn = ly.GraphObjects(i)
    if lgn is None:
        return False
    L, T, W, H = layer_rect(op, 0)
    for _ in range(3):
        lgn.Left = int(L + W - pad_r - lgn.Width)
        lgn.Top = int(T + pad_t)
        tb = op._true_bbox("Legend")
        if tb is None:
            return True
        dx = (L + W - pad_r) - tb[2]
        dy = (T + pad_t) - tb[1]
        if abs(dx) <= 2 and abs(dy) <= 2:
            return True
        lgn.Left = int(lgn.Left + dx)
        lgn.Top = int(lgn.Top + dy)
    return True


def clear_layer_titles(op, k):
    """清掉图层 k 的四个内置轴标题（顶/右轴会显示列名）。"""
    op.app.Execute('page.active = %d;' % (k + 1))
    op.app.Execute('xt.text$=""; yr.text$="";')


def hide_layer2_extras(op, graph=None):
    """图层2 画完曲线会自动长出图例和轴标题，全部关掉。"""
    if graph:
        op.app.Execute('win -a "%s";' % graph)
    op.app.Execute('page.active = 2;')
    op.app.Execute('legend.show = 0;')
    op.app.Execute('xb.text$=""; yl.text$=""; xt.text$=""; yr.text$="";')
    op.app.Execute('layer.x.showLabels=0; layer.x2.showLabels=0;')


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("data")
    ap.add_argument("--cards", nargs="+", default=["Cu2O", "Cu"])
    ap.add_argument("--x", nargs=2, type=float, default=None,
                    help="X 范围；不给就用数据实测范围")
    ap.add_argument("--name", default=None)
    ap.add_argument("--smooth", type=int, default=0)
    ap.add_argument("--card-y", type=float, default=CARD_Y_MAX)
    ap.add_argument("--legend-cards", action="store_true")
    ap.add_argument("--no-labels", action="store_true", help="不画卡片说明和晶面标注")
    ap.add_argument("--peaks", type=int, default=3, help="标几个最强峰的晶面（默认 3）")
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args(argv)

    src = os.path.abspath(a.data)
    stem = os.path.splitext(os.path.basename(src))[0]
    out_name = a.name or stem

    x, y = load_xy(src)
    if a.x is None:
        # 用户定的规矩：X 范围跟实测范围走，只有明确要求时才改
        xlim = (round(float(x.min()), 1), round(float(x.max()), 1))
    else:
        xlim = (a.x[0], a.x[1])
    print("X 范围: %.1f - %.1f（%s）" % (xlim[0], xlim[1],
                                     "实测范围" if a.x is None else "用户指定"))
    if a.smooth and a.smooth >= 5:
        from scipy.signal import savgol_filter
        y = savgol_filter(y, a.smooth | 1, 3)
        print("平滑: SG 窗口 %d 点" % (a.smooth | 1))
    cards = [load_card(c) for c in a.cards]
    for c in cards:
        if c.get("source") != "jade":
            print("!! 卡片 %s (%s) 未核对 JADE 导出" % (c["formula"], c["card"]))

    s_rows, ylim = build_sample(x, y, xlim)
    c_rows = build_cards(cards, xlim, base=-0.02 * a.card_y)
    print("数据:", src, " 样品行:", len(s_rows), " 卡片行:", len(c_rows))

    op = OriginPlot()
    for p in list(op.app.Pages):
        if p.Name != KEEP_PAGE:
            try:
                op.delete_page(p.Name)
            except Exception:
                pass

    # ---- 图层1：样品 ----
    book_s = op.new_book(s_rows, long_names=["2Theta", stem])
    graph = op.line_plot(book_s, ncols=1, colors=[SAMPLE_COLOR], widths=[SAMPLE_W])
    op.app.Execute('win -a "%s"; page.active = 1;' % graph)
    op.app.Execute('layer.x.from=%g; layer.x.to=%g; layer.y.from=%g; layer.y.to=%g;'
                   % (xlim[0], xlim[1], ylim[0], ylim[1]))

    op.app.Execute('layer.x.inc=10; layer.x.minorTicks=1;')
    op.app.Execute('layer.x.showAxes=3; layer.y.showAxes=3;')
    op.app.Execute('layer.y.tickLength=0; layer.y2.tickLength=0; layer.x2.tickLength=0;')
    op.app.Execute('layer.y.showLabels=0; layer.y2.showLabels=0;')
    op.app.Execute('layer.x.label.pt=22; layer.y.label.pt=22;')
    op.app.Execute('layer.x.label.font=1; layer.y.label.font=1;')
    op.app.Execute('xb.text$ = "2\u03b8 (degree)"; xb.fsize = 22;')
    op.app.Execute('yl.text$ = "Intensity (a.u.)"; yl.fsize = 22;')
    op.app.Execute('xt.text$=""; yr.text$="";')

    # ---- 图层2：卡片 ----
    op.app.Execute('win -a "%s"; layadd type:=1;' % graph)
    book_c = op.new_book(c_rows, long_names=["2Theta"] + [c["formula"] for c in cards])
    op.app.Execute('win -a "%s";' % graph)      # new_book 会把活动页切成工作簿
    for j in range(2, len(cards) + 2):
        op.app.Execute('plotxy iy:=[%s]1!(1,%d) plot:=200 ogl:=[%s]2!;'
                       % (book_c, j, graph))
    hide_layer2_extras(op, graph)
    # 卡片层是 layadd 建出来的**子图层**，坐标相对父层 —— 必须 0/0/100/100 才重合；
    # 照抄父层的 8.39/3.49/86.73/79.62 会再缩进一圈（实测框缩到 15.5%/63.8%）
    op.app.Execute('page.active = 2;')
    op.app.Execute('layer.left=0; layer.top=0; layer.width=100; layer.height=100;')
    op.app.Execute('layer.x.link = 1;')   # X 跟父层联动（缩放/平移同步）
    op.app.Execute('layer.y.link = 0;')   # Y 独立（0..CARD_Y_MAX），棒高就靠它调
    op.app.Execute('layer.x.from=%g; layer.x.to=%g; layer.x.inc=10;'
                   % (xlim[0], xlim[1]))
    op.app.Execute('layer.y.from=0; layer.y.to=%g;' % a.card_y)
    op.app.Execute('layer.x.showAxes=0; layer.y.showAxes=0;')
    for k in range(len(cards)):
        op.app.Execute('layer.plot%d.color = %d;' % (k + 1, _rgb(CARD_COLORS[k])))
        op.app.Execute('layer.plot%d.line.width = %g;' % (k + 1, STICK_W / 500.0))
    op.app.Execute('doc -u;')

    # 收尾：再确认一次卡片层 Y（父层 Y 改动不会带跑它，因为 link=0）
    op.app.Execute('page.active = 2;')
    op.app.Execute('layer.y.link = 0; layer.y.from=0; layer.y.to=%g;' % a.card_y)
    op.app.Execute('page.active = 1;')
    op.app.Execute('layer.y.from=%g; layer.y.to=%g;' % ylim)
    op.app.Execute('layer.x.showLabels = 1;')

    # ---- 图例 / Y 标题 ----
    lines = [r"\l(1) " + stem]
    if a.legend_cards:
        for i, c in enumerate(cards, start=2):
            lines.append(r"\l(%d) %s" % (i, c.get("label", c["formula"])))
    op.app.Execute('page.active = 1;')
    op.set_legend("\r\n".join(lines), size=22)
    op.app.Execute('legend.symbolwidth = 300;')   # 样例线缩短，图例变窄
    place_legend(op)

    # 固定页面 + 固定图层矩形：比 trim 稳，Y 标题/图例坐标都可直接算
    op.app.Execute('page.width=%g; page.height=%g;'
                   % (PAGE_W * 600, PAGE_H * 600))   # page.width 单位是 1/600 英寸
    set_rect(op, 0, LAYER[0] / 100 * PAGE_W * 600, LAYER[1] / 100 * PAGE_H * 600,
             LAYER[2] / 100 * PAGE_W * 600, LAYER[3] / 100 * PAGE_H * 600,
             PAGE_W * 600, PAGE_H * 600)
    for k in range(1, op.app.ActivePage.Layers.Count):
        op.app.Execute('page.active = %d;' % (k + 1))
        op.app.Execute('layer.left=0; layer.top=0; layer.width=100; layer.height=100;')
    print("页面: %.2f x %.2f 英寸" % (PAGE_W, PAGE_H))

    place_legend(op)
    names = ["Legend"] + ["cardlab%d" % (k + 1) for k in range(len(cards))]
    names += ["pk%d" % (k + 1) for k in range(a.peaks)]
    ok, report = op.check_overlap(names=tuple(names))
    for nm, n, verdict in report:
        print("  自检 %-10s 深色像素 %5d  %s" % (nm, n, verdict))

    # Y 轴标题拉开到 raman-plot 的 98 设备单位
    place_ytitle_gap(op, gap=98)
    # 右上角卡片说明 + 最强三个峰的晶面
    if not a.no_labels:
        add_card_labels(op, cards, CARD_COLORS[:len(cards)],
                        rows=s_rows, xlim=xlim, card_y_max=a.card_y)
        add_peak_labels(op, s_rows, cards, CARD_COLORS[:len(cards)], n=a.peaks)

    # 视觉自检：所有文字对象都不能压在曲线/轴线上
    names = ["Legend"] + ["cardlab%d" % (k + 1) for k in range(len(cards))]
    names += ["pk%d" % (k + 1) for k in range(a.peaks)]
    ok, report = op.check_overlap(names=tuple(names))
    for nm, n, verdict in report:
        print("  自检 %-10s 深色像素 %5d  %s" % (nm, n, verdict))
    if not ok:
        print("!! 有元素压线：调 --card-y / 字号，或手动在 Origin 里挪")

    files, cm, px = op.export_dpi(HERE, name=out_name, dpi=DPI, types=("tif",))
    print("导出: %.2f x %.2f cm -> %d x %d px" % (cm[0], cm[1], px[0], px[1]))
    for f, ok2, sz in files:
        print("  ", f, ok2, sz)

    proj = os.path.join(HERE, out_name + ".opju")
    if os.path.exists(proj) and not a.force:
        print("工程已存在，跳过保存:", proj)
    else:
        ok3 = False
        for _ in range(3):
            ok3, _e = op.save_project(proj, overwrite=True)
            if ok3:
                break
            time.sleep(0.8)
        print("工程:", proj, ok3)
    print("book:", book_s, "/", book_c, " graph:", graph)
    return 0


def _rgb(spec):
    r, g, b = [int(v) for v in spec.split(",")[:3]]
    return 0x01000000 | (b << 16) | (g << 8) | r


if __name__ == "__main__":
    raise SystemExit(main())
