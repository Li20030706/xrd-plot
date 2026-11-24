# -*- coding: utf-8 -*-
"""XRD 寻峰（绘图前的第一步）：读数据 -> 扣背景 -> 平滑 -> 找峰 -> 打表

用法：
    python xrd_peaks.py <数据文件...> [--min-rel 1.0] [--json out.json]

支持 Rigaku SmartLab 导出的 .txt（`;RAW4.00` 头 + `[Data]` 段）和任意两列文本。

寻峰口径（固定，别乱改）：
1. rolling-ball 背景：1.6° 窗口的最小值滤波再最大值滤波，SG 平滑
   —— 宽包（无定形/碳 002）会被当成背景扣掉，只留尖锐衍射峰
2. 残差 SG 平滑 0.15°
3. scipy find_peaks：prominence >= max(6σ, 满量程 1%)，FWHM <= 1.2°
   σ 取 sqrt(中位背景计数)，即泊松噪声
4. 相对强度 I/I0 = 扣背景峰高 / 最强峰高 × 100
"""
import argparse
import json
import os
import sys

import numpy as np
from scipy.ndimage import maximum_filter1d, minimum_filter1d
from scipy.signal import find_peaks, savgol_filter

LAMBDA = 1.5406  # Cu Kα1, Å

try:  # Windows 控制台默认 GBK，打印 Å/2θ 会炸
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def load_xy(path):
    """读两列 (2theta, intensity)，自动跳过文件头。"""
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


def rolling_ball(y, win):
    bg = minimum_filter1d(y, size=win, mode="nearest")
    bg = maximum_filter1d(bg, size=win, mode="nearest")
    return savgol_filter(bg, max(5, win | 1), 2)


def analyze(path, ball_deg=1.6, sg_deg=0.15, prom_sigma=6.0, max_fwhm=1.2,
            min_rel=1.0):
    x, y = load_xy(path)
    step = float(np.median(np.diff(x)))
    win = max(5, int(round(ball_deg / step)) | 1)
    bg = rolling_ball(y, win)
    r = np.clip(y - bg, 0, None)
    sgw = max(5, int(round(sg_deg / step)) | 1)
    rs = savgol_filter(r, sgw, 3)
    sigma = float(np.sqrt(np.median(np.maximum(bg, 1))))
    ymax = float(rs.max()) if rs.max() > 0 else 1.0
    prom = max(prom_sigma * sigma, 0.01 * ymax)
    idx, props = find_peaks(rs, prominence=prom, width=2)
    peaks = []
    for k, i in enumerate(idx):
        fw = float(props["widths"][k]) * step
        if fw > max_fwhm:
            continue
        h = float(rs[i])
        rel = h / ymax * 100.0
        if rel < min_rel:
            continue
        d = LAMBDA / (2.0 * np.sin(np.radians(x[i] / 2.0)))
        peaks.append({"two_theta": round(float(x[i]), 3), "d": round(d, 4),
                      "height": round(h, 1), "rel": round(rel, 1),
                      "fwhm": round(fw, 3)})
    peaks.sort(key=lambda p: -p["rel"])
    for n, p in enumerate(peaks, 1):
        p["no"] = n
    return {"file": os.path.basename(path), "n_points": int(len(x)),
            "range": [round(float(x[0]), 2), round(float(x[-1]), 2)],
            "step": round(step, 4), "sigma": round(sigma, 2),
            "threshold": round(prom, 1), "peaks": peaks}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--min-rel", type=float, default=1.0)
    ap.add_argument("--json")
    a = ap.parse_args(argv)
    out = []
    for f in a.files:
        r = analyze(f, min_rel=a.min_rel)
        out.append(r)
        print("=" * 72)
        print("%s  %d 点  %.2f-%.2f°  step %.4f  σ≈%.1f  阈值≈%.1f"
              % (r["file"], r["n_points"], r["range"][0], r["range"][1],
                 r["step"], r["sigma"], r["threshold"]))
        print("%-4s %8s %7s %8s %7s %7s" % ("#", "2θ(°)", "d(Å)", "高度", "I/I0", "FWHM"))
        for p in r["peaks"]:
            print("%-4d %8.3f %7.3f %8.1f %7.1f %7.3f"
                  % (p["no"], p["two_theta"], p["d"], p["height"], p["rel"], p["fwhm"]))
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump(out, fh, ensure_ascii=False, indent=2)
        print("已写:", a.json)
    return 0


if __name__ == "__main__":
    sys.exit(main())
