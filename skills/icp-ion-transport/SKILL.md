---
name: icp-ion-transport
description: Agilent ICP-OES 膜离子传输 / 价选择性数据处理管线（replicate-first）。读 1_Original Data 或原始 CSV → 加法累计采样校正 → 补出 8-sheet Results.xlsx（浓度/通量/选择性/质量平衡/电流效率）→ 画 Figure 1(浓度)/2(通量)/3(选择性)。用于处理掺杂系列(doping)或电流密度(current density)的 CEM/AEM 离子迁移实验、计算 Ca²⁺/Na⁺ permselectivity、合并平行批次、重算或重画图。
---

# ICP-OES 离子传输管线 (icp-ion-transport)

核心模块 `replicate_selectivity_workflow.py` 被各 `calculate_results_*.py` 与
`plot_figure3_*.py` 以 `import` 调用。同目录 `USAGE.md` 是详细版说明。

## 目录结构假设
脚本用 `Path(__file__).resolve().parents[3]` 定位核心模块。还原到原目录结构即可直接跑：
```
<父>/replicate_selectivity_workflow.py
<父>/<研究>/Figure/Results/calculate_results.py
<父>/<研究>/Figure/Figure N - .../plot_figureN.py
```

## 运行顺序 (process → calculate → plot)
1. **process**（原始 CSV → `1_Original Data` + `Calibration` 两 sheet）
   - 合并电流密度：`process_merge_current_density.py`（读两批 CSV，Ca 317.933，池化重复孔，`DROP_GROUPS` 可整组剔除）。
   - 掺杂系列：无 process 脚本，直接复用既有已清洗的 `1_Original Data`（Ca 422.673）。
2. **calculate**（补出 sheet 2–8）
   - `calculate_results_current_density.py` / `calculate_results_doping.py`，调用
     `rebuild_results_workbook(xlsx, default_current_density=, time_points=[0,120], exclude=)`。
3. **plot**：`plot_figure1.py`（浓度）、`plot_figure2.py`（通量）、`plot_figure3_*.py`（选择性）。

## 运行命令
有 uv：
```
uv run --with openpyxl,scipy,numpy,matplotlib python <脚本>
```
无 uv（如本机 Windows，python 已自带科学栈）：
```
python <脚本>
```

## 方法学要点（务必记住）
- **采样稀释校正 = 加法累计**（不是几何式！）：`C_corr(t_i)=C_meas(t_i)+(Vs/V)·Σ_{j<i}C_meas(t_j)`，按真实采样次序累计。
- **replicate-first**：每个重复孔独立算到底（浓度/通量/逐孔 selectivity），Mean = 有效孔平均；支持任意孔数。
- **flux**：`J = dC_corr · V / A / t`（mmol·cm⁻²·min⁻¹），per replicate 后再平均。
- **端点分析**：`time_points=[0,120]` 只导出 0/120，但加法校正仍用中间测点 → 120min 保留真实 i=4 历史。
- **选择性 NQ 过滤**：某孔 JCa_eff≤0 / JNa_eff≤0 / |JNa_eff|<1e-6 记 NQ，不计入 mean；DIL 侧用 −J。
- **误差棒**：各孔最终量取总体标准差 std(ddof=0)，无解析误差传递。
- 参数：V=20mL, Vs=0.4mL, A=1.54cm², 稀释75×, M_Ca=40.08, M_Na=22.99, F=96485.33212。

## 两个关键开关
- `rebuild_results_workbook(..., exclude={(group, time)})`：把某 (组,时间) 当缺失，从加法校正剔除、不导出。
  例：`{("10wt% 10mA*cm^-2", 90)}`（已确认的 2 倍稀释坏点）。
- `process_merge_current_density.py` 里 `DROP_GROUPS`：整组剔除，源头不留记录。
  例：`{"5wt% 1mA*cm^-2","10wt% 1mA*cm^-2"}`（1mA 扩散主导，已删）。
- `rebuild_results_workbook(..., vs_overrides={(group, rep): Vs_mL})`：让某些重复孔用**各自的采样体积**做加法累计校正（默认全用全局 V_sample_mL）。
  用于合并不同采样条件的孔，例：old membrane（Vs=0.1）与并入的 OLD 5wt%（Vs=0.4）作平行重复时各用各的 Vs。
  模块还兼容「某组缺 DIL 侧」（只出 CON）与 Figure 3 doping 模式里的**非 wt% 参比组**（如 FKS-PET-130，灰色柱排最右）。

## Figure 3 模式
`plot_selectivity_figures(xlsx, out_dir, mode=...)`：
- `mode="current_density"`：每个 wt 一张图，selectivity vs 电流密度。
- `mode="doping"`：柱状图，**横轴 = Resin loading (wt%)，按 loading 从小到大排列**（0→1→3→5→10），
  柱色用**顺序蓝渐变**（随 loading 加深，`plt.cm.Blues` 0.30→0.92，细灰边框）以呼应有序变量。
  默认**输出两张**：CON-based = `Figure3_Selectivity.png`，DIL-based = `Figure3_Selectivity_DIL.png`
  （两个区块的 selectivity 一直都在 `6_Selectivity` sheet 里算好，CON/DIL 各一套；DIL 侧用 −J 的迁出量）。
  用 `plot_selectivity_figures(..., doping_bases=("CON",))` 可**只画 CON**（如 old membrane 多数组没测 DI 室浓度，只出 concentrate 室选择性；DIL 数值仍保留在 sheet 里，只是不画）。
- **时间分辨选择性**：`plot_selectivity_timecourse(xlsx, out_dir, vs_overrides=, exclude=)` → 双 panel（上 CON、下 DIL）的 selectivity vs time 折线图。它**直接从 `1_Original Data` 在所有测点重算**逐孔选择性，**不依赖端点导出**（所以端点柱状图设端点几都不影响它）。输出 `Figure3_Selectivity_timecourse.png`。

## 保真自检（换机后建议复跑）
用 `rebuild_results_workbook(既有Results, time_points=None)`（不传 exclude）对原始 Results 复跑，
sheet 2–8 应与原值 **0 diff**（证明方法学未漂移）。
