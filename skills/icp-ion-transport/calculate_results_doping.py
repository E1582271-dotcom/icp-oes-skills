# -*- coding: utf-8 -*-
"""Rebuild Results.xlsx (sheets 2-8) for the New-membrane doping series.

Input 1_Original Data + Calibration are reused from the original (already
cleaned / calibrated on Ca 422.673).  Endpoint-only analysis: only 0 and 120
min are exported; the additive sampling correction still uses the measured
0/30/60/90 points so the 120 min point keeps its real i=4 dilution history.
All groups were run at 10 mA*cm^-2.
"""
from pathlib import Path
import sys

PERFORMANCE_DIR = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PERFORMANCE_DIR))

from replicate_selectivity_workflow import rebuild_results_workbook

if __name__ == "__main__":
    result = rebuild_results_workbook(
        Path(__file__).with_name("Results.xlsx"),
        default_current_density=10.0,
        time_points=[0, 120],
    )
    print(f"Saved: {result['workbook']}")
    print(f"Groups: {result['groups']}")
    print(f"Time points: {result['time_points']}")
