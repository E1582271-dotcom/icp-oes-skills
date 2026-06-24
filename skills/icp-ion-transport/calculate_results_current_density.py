# -*- coding: utf-8 -*-
"""Rebuild Results.xlsx (sheets 2-8) for the merged current-density study.

Endpoint-only analysis: only the 0 and 120 min rows are exported, but the
additive sampling correction still uses every measured intermediate point
(0/30/60/90 -> the 120 min point keeps its real i=4 dilution history).
"""
from pathlib import Path
import sys

PERFORMANCE_DIR = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PERFORMANCE_DIR))

from replicate_selectivity_workflow import rebuild_results_workbook

if __name__ == "__main__":
    result = rebuild_results_workbook(
        Path(__file__).with_name("Results.xlsx"),
        default_current_density=None,   # parsed from group names (1/5/10 mA*cm^-2)
        time_points=[0, 120],
        # Data-check finding: 10wt% 10mA t=90 is a confirmed dilution-error batch
        # (all 4 cells, both ions ~0.57x neighbours). Drop it as missing so it is
        # skipped from the additive correction; raw stays in 1_Original Data.
        exclude={("10wt% 10mA*cm^-2", 90)},
    )
    print(f"Saved: {result['workbook']}")
    print(f"Groups: {result['groups']}")
    print(f"Time points: {result['time_points']}")
