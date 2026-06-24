# -*- coding: utf-8 -*-
"""Figure 2: Ion flux (CON + DI) vs time, mean +/- replicate std.

Reads the replicate-first Results.xlsx (sheets 4 & 5).  Endpoint-only analysis
(0 & 120 min) means each series is two points.  Mode auto-detected as in
plot_figure1.py.
"""
import os
import re

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.lines import Line2D
import numpy as np
import openpyxl

plt.rcParams.update({
    "font.family": "sans-serif", "font.sans-serif": ["Arial", "DejaVu Sans"],
    "mathtext.fontset": "stixsans", "font.size": 8, "axes.linewidth": 0.8,
    "axes.labelsize": 9, "xtick.direction": "in", "ytick.direction": "in",
    "lines.linewidth": 1.2, "lines.markersize": 5, "legend.frameon": False,
    "figure.dpi": 300, "savefig.dpi": 600, "pdf.fonttype": 42, "ps.fonttype": 42,
})

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "..", "Results", "Results.xlsx")
EM = "—"
_NPG = ["#E64B35", "#4DBBD5", "#00A087", "#3C5488", "#F39B7F", "#8491B4"]
_MKRS = ["o", "s", "p", "D", "^", "v", "h", "X"]


def _f(v):
    if v is None or v == EM or isinstance(v, str):
        return np.nan
    return float(v)


def parse_header(h):
    """Robust to 'Metric\\nRep-k' and 'Metric Rep-k\\n(unit)' header styles."""
    s = str(h).replace("\n", " ")
    s = re.sub(r"\(.*?\)", "", s)          # drop unit
    if "Mean" in s:
        return s.split("Mean")[0].strip(), "mean"
    if "Rep-" in s:
        return s.split("Rep-")[0].strip(), "rep"
    return None


def read_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    band, header = rows[3], rows[4]
    starts = [(ci, str(v)) for ci, v in enumerate(band) if ci >= 1 and v not in (None, "")]
    bounds = [s[0] for s in starts] + [len(header)]
    times, drows = [], []
    for r in rows[5:]:
        if r and isinstance(r[0], (int, float)):
            times.append(int(r[0])); drows.append(r)
    data, groups = {}, []
    for gi, (cs, name) in enumerate(starts):
        end = bounds[gi + 1]; groups.append(name)
        mcols = {}
        for ci in range(cs, end):
            h = header[ci]
            if not h:
                continue
            parsed = parse_header(h)
            if not parsed:
                continue
            metric, tag = parsed
            mcols.setdefault(metric, {"reps": [], "mean": None})
            if tag == "mean":
                mcols[metric]["mean"] = ci
            elif tag == "rep":
                mcols[metric]["reps"].append(ci)
        d = {}
        for metric, cols in mcols.items():
            means = np.array([_f(dr[cols["mean"]]) if cols["mean"] is not None else np.nan for dr in drows])
            stds = []
            for dr in drows:
                reps = [_f(dr[c]) for c in cols["reps"]]
                reps = [x for x in reps if not np.isnan(x)]
                stds.append(np.std(reps, ddof=0) if len(reps) > 1 else 0.0)
            d[metric] = (means, np.array(stds))
        data[name] = d
    return groups, times, data


def split_name(name):
    m = re.match(r"\s*(\S+wt%)\s+(.*\S)\s*$", name)
    return (m.group(1), m.group(2).strip()) if m else (name, "")


def regline(x, y):
    mask = np.isfinite(x) & np.isfinite(y)
    if mask.sum() < 2:
        return None, None
    s, b = np.polyfit(x[mask], y[mask], 1)
    xf = np.linspace(x[mask].min(), x[mask].max(), 50)
    return xf, s * xf + b


wb = openpyxl.load_workbook(XLSX, data_only=True)
groups, times, con = read_sheet(wb["4_CON Ion Flux"])
_, _, di = read_sheet(wb["5_DI Ion Flux"])
T = np.array(times)
SCALE = 1e3
JCA, JNA = "J Ca2+", "J Na+"

is_cd = any(re.search(r"\d+\s*mA", g) for g in groups)
if is_cd:
    fractions = list(dict.fromkeys(split_name(g)[0] for g in groups))
    members = {fr: [g for g in groups if split_name(g)[0] == fr] for fr in fractions}
    currents = list(dict.fromkeys(split_name(g)[1] for g in groups))
    color = {c: _NPG[i % len(_NPG)] for i, c in enumerate(currents)}
    marker = {c: _MKRS[i % len(_MKRS)] for i, c in enumerate(currents)}
    key = lambda g: split_name(g)[1]
else:
    fractions = ["all"]; members = {"all": groups}
    color = {g: _NPG[i % len(_NPG)] for i, g in enumerate(groups)}
    marker = {g: _MKRS[i % len(_MKRS)] for i, g in enumerate(groups)}
    key = lambda g: g

UNIT = r"($10^{-3}$ mmol cm$^{-2}$ min$^{-1}$)"


def panel(ax, src, metric, ylabel, plabel, sel, legend=False):
    for g in sel:
        m, s = src[g][metric]
        kk = key(g); msk = ~np.isnan(m)
        ax.errorbar(T[msk], m[msk] * SCALE, yerr=s[msk] * SCALE, fmt=marker[kk], color=color[kk],
                    ms=5, capsize=2, capthick=0.8, label=kk, zorder=4)
        xf, yf = regline(T, m * SCALE)
        if xf is not None:
            ax.plot(xf, yf, color=color[kk], lw=1.0, ls="--", alpha=0.8, zorder=2)
    ax.axhline(0, color="#888888", lw=0.7, ls=":")
    ax.set_xlabel("Time (min)"); ax.set_ylabel(ylabel)
    ax.set_xlim(-8, max(times) + 15)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(30))
    ax.text(-0.12, 1.05, plabel, transform=ax.transAxes, fontsize=10, fontweight="bold", va="top")
    if legend:
        h = [Line2D([0], [0], color=color[key(g)], marker=marker[key(g)], ms=5, lw=1.0, ls="--",
                    label=key(g)) for g in sel]
        ax.legend(handles=h, loc="best", fontsize=6, ncol=2, handletextpad=0.3,
                  columnspacing=0.8, labelspacing=0.3)


def sync(*axes):
    lo = min(a.get_ylim()[0] for a in axes); hi = max(a.get_ylim()[1] for a in axes)
    for a in axes:
        a.set_ylim(lo, hi)


for fr in fractions:
    sel = members[fr]
    fig, ax = plt.subplots(2, 2, figsize=(7.5, 5.5))
    fig.subplots_adjust(left=0.10, right=0.97, top=0.90, bottom=0.10, hspace=0.35, wspace=0.35)
    if fr != "all":
        fig.suptitle(fr, fontsize=11, fontweight="bold", y=0.98)
    panel(ax[0, 0], con, JCA, r"$J_{\mathrm{Ca^{2+}}}$ (CON) " + UNIT, "a", sel, legend=True)
    panel(ax[0, 1], con, JNA, r"$J_{\mathrm{Na^+}}$ (CON) " + UNIT, "b", sel)
    panel(ax[1, 0], di, JCA, r"$J_{\mathrm{Ca^{2+}}}$ (DI) " + UNIT, "c", sel)
    panel(ax[1, 1], di, JNA, r"$J_{\mathrm{Na^+}}$ (DI) " + UNIT, "d", sel)
    sync(ax[0, 0], ax[0, 1]); sync(ax[1, 0], ax[1, 1])
    tag = fr.replace("%", "").replace(" ", "") if fr != "all" else "all"
    for fmt in ("png", "pdf"):
        out = os.path.join(HERE, f"Figure2_Flux_{tag}.{fmt}")
        fig.savefig(out, format=fmt, bbox_inches="tight")
        print(f"Saved: {out}")
    plt.close(fig)
