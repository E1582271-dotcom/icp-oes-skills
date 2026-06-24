# -*- coding: utf-8 -*-
"""Figure 1: Concentration (CON + DI) vs time, mean +/- replicate std.

Reads the replicate-first Results.xlsx (sheets 2 & 3, Rep-1..N / Mean layout).
Auto-detects mode:
  * current_density : group names like "5wt% 5mA*cm^-2" -> one figure per wt
                      fraction, one line per current density.
  * doping          : plain wt-fraction groups -> single overlay figure.
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
    "font.family": "sans-serif", "font.sans-serif": ["Arial", "Helvetica", "DejaVu Sans"],
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


def read_sheet(ws):
    """Return (groups, times, data) where data[group][metric] = (mean[], std[])."""
    rows = list(ws.iter_rows(values_only=True))
    band = rows[3]
    header = rows[4]
    starts = [(ci, str(v)) for ci, v in enumerate(band) if ci >= 1 and v not in (None, "")]
    bounds = [s[0] for s in starts] + [len(header)]
    # time rows
    times, drows = [], []
    for r in rows[5:]:
        if r and isinstance(r[0], (int, float)):
            times.append(int(r[0])); drows.append(r)
    data = {}
    groups = []
    for gi, (cs, name) in enumerate(starts):
        end = bounds[gi + 1]
        groups.append(name)
        metric_cols = {}
        for ci in range(cs, end):
            h = header[ci]
            if not h:
                continue
            parts = str(h).split("\n")
            metric = parts[0].strip()
            tag = parts[1].strip() if len(parts) > 1 else ""
            metric_cols.setdefault(metric, {"reps": [], "mean": None})
            if tag.lower().startswith("mean"):
                metric_cols[metric]["mean"] = ci
            elif tag.lower().startswith("rep"):
                metric_cols[metric]["reps"].append(ci)
        d = {}
        for metric, cols in metric_cols.items():
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
    if m:
        return m.group(1), m.group(2).strip()
    return name, ""


def regline(x, y):
    mask = np.isfinite(x) & np.isfinite(y)
    if mask.sum() < 2:
        return None, None
    s, b = np.polyfit(x[mask], y[mask], 1)
    xf = np.linspace(x[mask].min(), x[mask].max(), 50)
    return xf, s * xf + b


wb = openpyxl.load_workbook(XLSX, data_only=True)
groups, times, con = read_sheet(wb["2_CON Concentration"])
_, _, di = read_sheet(wb["3_DI Concentration"])
T = np.array(times)

is_cd = any(re.search(r"\d+\s*mA", g) for g in groups)
CA, NA = "Ca2+ corr", "Na+ corr"

if is_cd:
    fractions = list(dict.fromkeys(split_name(g)[0] for g in groups))
    members = {fr: [g for g in groups if split_name(g)[0] == fr] for fr in fractions}
    currents = list(dict.fromkeys(split_name(g)[1] for g in groups))
    color = {c: _NPG[i % len(_NPG)] for i, c in enumerate(currents)}
    marker = {c: _MKRS[i % len(_MKRS)] for i, c in enumerate(currents)}
    key = lambda g: split_name(g)[1]
else:
    fractions = ["all"]
    members = {"all": groups}
    color = {g: _NPG[i % len(_NPG)] for i, g in enumerate(groups)}
    marker = {g: _MKRS[i % len(_MKRS)] for i, g in enumerate(groups)}
    key = lambda g: g


def panel(ax, src, metric, ylabel, plabel, sel, legend=False):
    for g in sel:
        m, s = src[g][metric]
        kk = key(g)
        msk = ~np.isnan(m)
        ax.errorbar(T[msk], m[msk], yerr=s[msk], fmt=marker[kk], color=color[kk],
                    ms=5, capsize=2, capthick=0.8, mec="white", mew=0.35, label=kk, zorder=4)
        xf, yf = regline(T, m)
        if xf is not None:
            ax.plot(xf, yf, color=color[kk], lw=1.0, ls="--", alpha=0.8, zorder=2)
    ax.set_xlabel("Time (min)"); ax.set_ylabel(ylabel)
    ax.set_xlim(-8, max(times) + 15)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(30))
    ax.text(-0.12, 1.05, plabel, transform=ax.transAxes, fontsize=10, fontweight="bold", va="top")
    if legend:
        h = [Line2D([0], [0], color=color[key(g)], marker=marker[key(g)], ms=5, lw=1.0, ls="--",
                    mec="white", mew=0.35, label=key(g)) for g in sel]
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
    panel(ax[0, 0], con, CA, r"Ca$^{2+}$ (CON) (mmol L$^{-1}$)", "a", sel, legend=True)
    panel(ax[0, 1], con, NA, r"Na$^+$ (CON) (mmol L$^{-1}$)", "b", sel)
    panel(ax[1, 0], di, CA, r"Ca$^{2+}$ (DI) (mmol L$^{-1}$)", "c", sel)
    panel(ax[1, 1], di, NA, r"Na$^+$ (DI) (mmol L$^{-1}$)", "d", sel)
    sync(ax[0, 0], ax[0, 1]); sync(ax[1, 0], ax[1, 1])
    tag = fr.replace("%", "").replace(" ", "") if fr != "all" else "all"
    for fmt in ("png", "pdf"):
        out = os.path.join(HERE, f"Figure1_Concentration_{tag}.{fmt}")
        fig.savefig(out, format=fmt, bbox_inches="tight")
        print(f"Saved: {out}")
    plt.close(fig)
