# ICP-OES 离子传输管线 — 使用说明

核心模块：`replicate_selectivity_workflow.py`
被各 `calculate_results_*.py` 与 `plot_figure3_*.py` 以 `import` 调用。

## 运行顺序
1. **process**（原始 CSV → `1_Original Data` + `Calibration` 两 sheet）
   - 合并电流密度：`process_merge_current_density.py`（读两批 CSV，Ca 317.933，池化重复孔）。
   - 掺杂系列无 process 脚本：直接复用既有已清洗的 `1_Original Data`（Ca 422.673）。
2. **calculate**（补出 sheet 2–8）
   - `calculate_results_current_density.py` / `calculate_results_doping.py`，调用
     `rebuild_results_workbook(xlsx, default_current_density=, time_points=, exclude=)`。
3. **plot**：`plot_figure1.py`（浓度，读 sheet2/3）、`plot_figure2.py`（通量，读 sheet4/5）、
   `plot_figure3_*.py`（选择性，调用模块 `plot_selectivity_figures`）。

运行（每条都用 uv）：
```
uv run --with openpyxl,scipy,numpy,matplotlib python <脚本>
```
脚本用 `Path(__file__).resolve().parents[3]` 定位核心模块——放回原目录结构
（`<父>/replicate_selectivity_workflow.py` + `<父>/<研究>/Figure/Results/calculate_results.py`）即可直接跑。

## 方法学要点（务必记住）
- **采样稀释校正 = 加法累计**（不是几何式！旧 sheet 注释写错过）：
  `C_corr(t_i) = C_meas(t_i) + (Vs/V)·Σ_{j<i} C_meas(t_j)`，按真实采样次序累计。
- **replicate-first**：每个重复孔独立算到底（浓度/通量/逐孔 selectivity），Mean = 有效孔平均；
  支持任意孔数（单批 2 孔、两批池化 4 孔）。
- **端点分析**：`time_points=[0,120]` 只导出 0/120，但加法校正仍用中间测点 → 120min 保留真实 i=4 历史。
- **选择性 NQ 过滤**：某孔 JCa_eff≤0 / JNa_eff≤0 / |JNa_eff|<1e-6 记 NQ，不计入 mean。
- **误差棒**：对各孔最终量取总体标准差 std(ddof=0)，**无解析误差传递**；选择性误差 = 逐孔比值的 std。
- 参数：V=20mL, Vs=0.4mL, A=1.54cm², 稀释75×, M_Ca=40.08, M_Na=22.99, F=96485.33212。

## 两个关键开关
- `rebuild_results_workbook(..., exclude={(group, time)})`：把某 (组,时间) 当缺失，从加法校正剔除、不导出。
  当前用于 `{("10wt% 10mA*cm^-2", 90)}`（已确认的 2 倍稀释坏点）。
- `process_merge_current_density.py` 里 `DROP_GROUPS`：整组剔除，源头不留记录。
  当前 = `{"5wt% 1mA*cm^-2","10wt% 1mA*cm^-2"}`（1mA 扩散主导，已删）。

## 保真验证（换机后建议复跑一次自检）
用 `rebuild_results_workbook(既有Results, time_points=None)`（不传 exclude）对原始
Current Density / New membrane 的 Results 复跑，sheet 2–8 应与原值 **0 diff**（证明方法学未漂移）。
