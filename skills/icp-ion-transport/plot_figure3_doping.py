# -*- coding: utf-8 -*-
"""Figure 3: permselectivity per doping group (replicate-first, endpoint)."""
from pathlib import Path
import sys

PERFORMANCE_DIR = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PERFORMANCE_DIR))

from replicate_selectivity_workflow import plot_selectivity_figures

if __name__ == "__main__":
    here = Path(__file__).resolve().parent
    workbook = here.parent / "Results" / "Results.xlsx"
    for p in plot_selectivity_figures(workbook, here, mode="doping"):
        print(f"Saved: {p}")
