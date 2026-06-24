# ICP-OES Skills — 膜离子传输数据处理技能包

一套用于 **Agilent ICP-OES** 导出数据的 [Claude Code Skills](https://docs.claude.com/en/docs/claude-code/skills)，覆盖从原始 CSV 到**标准曲线标定 → 浓度计算 → 通量(flux) → 价选择性(permselectivity)**的完整流程。面向**电渗析(ED)/ 离子交换膜(CEM/AEM)离子迁移实验**。

> 方法学核心：采样稀释校正采用**加法累计**模型，并坚持 **replicate-first**（每个重复孔独立算到底，再取均值），保证通量与选择性的统计口径一致。

## 包含的 4 个 Skill

| Skill | 作用 | 用法 |
|---|---|---|
| [`icp-calibration`](./skills/icp-calibration) | 读 Agilent CSV，自动提取 STD 标准点，对 Ca 393.366 / Na 588.995 做分段线性拟合，输出标定图 + 可复现脚本 `plot_calibration.py` | `/icp-calibration <CSV路径>` |
| [`icp-process`](./skills/icp-process) | 单文件一键全流程：原始导出 → 拟合标准曲线 → 算样品浓度 → 输出分组着色的 `results.xlsx` | `/icp-process <原始文件路径>` |
| [`icp-merge`](./skills/icp-merge) | 扫描父目录下所有日期子文件夹，各自用当天标定曲线处理，合并输出 `merged_results.xlsx` | `/icp-merge <父目录路径>` |
| [`icp-ion-transport`](./skills/icp-ion-transport) | **通量 + 选择性主力管线**：补出 8-sheet `Results.xlsx`（浓度/通量/选择性/质量平衡/电流效率），并画 Figure 1–3 | 见下方 |

## icp-ion-transport 管线

```
process  →  calculate  →  plot
```

1. **process**：原始 CSV → `1_Original Data` + `Calibration` 两张 sheet
2. **calculate**：`rebuild_results_workbook(...)` 补出 sheet 2–8
3. **plot**：`plot_figure1.py`(浓度) / `plot_figure2.py`(通量) / `plot_figure3_*.py`(选择性)

运行（有 [uv](https://github.com/astral-sh/uv)）：

```bash
uv run --with openpyxl,scipy,numpy,matplotlib python <脚本>
```

无 uv（科学栈已自带）：

```bash
python <脚本>
```

### 方法学要点

- **采样稀释校正 = 加法累计**：`C_corr(t_i) = C_meas(t_i) + (Vs/V)·Σ_{j<i} C_meas(t_j)`，按真实采样次序累计。
- **replicate-first**：每个重复孔独立算到浓度/通量/逐孔 selectivity，Mean = 有效孔平均，支持任意孔数。
- **flux**：`J = dC_corr · V / A / t`（mmol·cm⁻²·min⁻¹），先 per replicate 再平均。
- **选择性 NQ 过滤**：某孔 `JCa_eff≤0` / `JNa_eff≤0` / `|JNa_eff|<1e-6` 记 NQ，不计入均值。
- 默认参数：`V=20 mL, Vs=0.4 mL, A=1.54 cm², 稀释 75×, M_Ca=40.08, M_Na=22.99, F=96485.33212`。

详细说明见 [`skills/icp-ion-transport/USAGE.md`](./skills/icp-ion-transport/USAGE.md)。

## 安装为 Claude Code Skills

把需要的文件夹复制到你的 skills 目录即可：

```bash
# 全局（对所有项目可用）
git clone https://github.com/E1582271-dotcom/icp-oes-skills.git
cp -r icp-oes-skills/skills/* ~/.claude/skills/
```

之后在 Claude Code 里用 `/icp-calibration`、`/icp-process`、`/icp-merge` 调用；`icp-ion-transport` 通过描述自动触发或手动调用其脚本。

## 依赖

Python 3.9+，`numpy` / `scipy` / `matplotlib` / `openpyxl`。

## License

[MIT](./LICENSE)
