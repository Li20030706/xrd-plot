---
name: xrd-plot
description: Draw stacked XRD figures in Origin (sample pattern on top, PDF/JCPDS reference sticks below) via COM, including peak finding, card import and visual self-check.
---

# XRD stacked figure (Origin COM)

A fixed, opinionated layout. Follow the steps; do not redesign the style.

All scripts live in this repository. Origin must already be running (OriginPro 2021+,
your own visible window) — the scripts attach to it, they never launch a headless copy.

## Workflow (do not skip steps)

### Step 0 — ask before drawing

1. **How many peaks should get an hkl label?** (default 3)
2. Which file to plot

The X range follows the **measured range of the file**. Only override it when the user
explicitly asks for a range (e.g. `--x 20 70`). If the low-angle part is obviously the
direct-beam tail (intensity far above the rest of the pattern), say so and suggest
starting at 10–20°.

### Step 1 — find the peaks

```bash
python xrd_peaks.py data.txt
```

Prints 2θ / d / I/I0 / FWHM.

- Peaks below ~10° are usually beam-stop or holder artefacts — do not report them as phases
- Amorphous samples (broad humps, no sharp peaks) must be reported as "no crystalline phase"

### Step 2 — propose cards, wait for confirmation

Match the peak list against `xrd_cards.json` and show the user a table: 2θ, d, I/I0,
candidate assignment, and the evidence (position offset, intensity ratio).

**Cards must come from a JADE export or a public structure, never from memory.**

```bash
python xrd_cards.py --list
python xrd_cards.py --import "/path/to/PDF#05-0667.txt" --key Cu2O
python xrd_cards.py --search "cuprite"                  # free: search COD
python xrd_cards.py --from-cod 9007497 --key Cu2O_cod   # computed pattern (comparison only)
```

Hand-typed cards are marked `source: memory` and trigger a warning in `xrd_plot.py`.
Only draw after the user confirms the phases.

### Step 3 — draw

```bash
python xrd_plot.py data.txt --cards Cu2O Cu --smooth 15 --peaks 3
```

- `--smooth N` — Savitzky-Golay window in points; 15 (≈0.3°) suits most XRD data, keep ≤21
- `--card-y V` — card layer Y range; 1000 makes the sticks 10 % of the frame height
- `--no-labels` — no captions at all

### Step 4 — visual self-check (mandatory)

The script prints one line per text object:

```
check Legend     dark px     0  clean
check cardlab1   dark px     0  clean
check pk1        dark px     0  clean
```

**Any non-zero count means an element sits on the curve or the axes** — move it before
exporting. Then open the TIF yourself and look at it:

```python
from PIL import Image
Image.open("out.tif").convert("RGB").save("_check.png")
```

Check: labels clipped by the frame, distance between peak labels and peak tips, card
captions touching the curve, legend running past the right frame.

### Step 5 — ask whether the user is happy

1. Never declare it finished on your own
2. Happy → deliver the `.tif` and `.opju`
3. Not happy:
   - they describe the change → adjust the parameters and re-run
   - they want to drag things themselves → bring the Origin window to the front
     (`win -a "Graph1"` + `win32gui.SetForegroundWindow`), then read their edits back with
     `python xrd_snapshot.py --diff base.json` and fold the values into the script

## Fixed style (confirmed with the user — do not change)

| Item | Value |
|---|---|
| Layers | 2: sample on layer 1, cards on a child layer (rect `0/0/100/100`) |
| Page | fixed 8.43 × 7.40 in; layer rect 8.39/3.49/86.73/79.62 |
| Sample Y | normalised 0–100 placed at 28–128, axis 0–140 |
| Card Y | 0–1000, card intensity = I/I0 (max 100) → stick height = 10 % of frame |
| Card layer links | `x.link=1` (follow parent X), `y.link=0` (independent Y) |
| Frame | all four axes; no ticks on top/right; no Y ticks or numbers; X major 10, minor 5 |
| Axis titles | `2θ (degree)` / `Intensity (a.u.)`, Arial 22; Y title right edge 98 units from the axis |
| Legend | top-right, Arial 22, no frame, sample name only; right pad 177, top pad 318, `legend.symbolwidth=300` |
| Sample line | `#8E9CFF`, 2.0 pt |
| Card sticks | `#FDC1C1` / `#CEF297`, 3.0 pt, sitting on the X axis |
| Card captions | right side of the blank band between curve and sticks, right-aligned, Arial **20**, phase colour |
| Peak labels | hkl of the N strongest peaks, phase colour, 30 units above the peak |
| Deliverables | 300 dpi `.tif` (2529×2220) + `.opju`, same stem as the input file |

## Pitfalls (all hit in practice)

**Layers**
- `layadd type:=1` creates a **child layer whose coordinates are relative to the parent**:
  set `layer.left=0; top=0; width=100; height=100` or it shrinks one step and the sticks float
- the new layer defaults to `y.link=1` (follows the parent Y) — a 0–140 parent squashes the
  sticks to 71 % of the frame and they cross the curve; set `y.link=0`
- `layerN.xxx` prefixes are unreliable (they hit the active layer): use `page.active = N` or
  `Layers(k).Activate()`
- `new_book()` switches the active page to the workbook — `win -a "<graph>"` afterwards
- the 20–70 numbers at the top are the **card layer's** x labels; turn those off

**Curve styling**
- `plotxy iy:=(1,2:N)` groups the plots and the line width can only be set for the whole
  group → add curves one at a time with `plotxy … ogl:=[G]2!`
- `set %C -cl/-w` does not advance the way you expect on multi-curve layers → set colours
  with `DataPlot.SetNumProp("color", 0x01BBGGRR)`

**Sticks**
- **do not rely on NaN breaks**: Origin connects across missing values and draws diagonals
  between sticks → emit `base → top → base` triples, with the base below the axis minimum

**Text**
- the built-in Y axis title only moves through `yl.x` (at `yl.x≈0` it is off-page; on an
  8.43×7.40 page only 16–19 puts it in the left margin) → `place_ytitle_gap` scans that window
- a rotated label's `obj.Width` is stale until a redraw, and `_true_bbox` mis-measures rotated
  objects → do not use a rotated label as the axis title
- a wide legend runs into the tallest peak → `legend.symbolwidth=300`

**Saving**
- `.opju` needs a backslash absolute path; `app.Save` occasionally returns False right after an
  export (file handle not released) → retry three times
