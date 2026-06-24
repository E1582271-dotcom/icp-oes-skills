---
name: icp-calibration
description: 读取 Agilent ICP-OES 导出的 CSV 文件，自动提取 STD 标准点，对 Ca 393.366 nm 和 Na 588.995 nm 做分段线性拟合，输出可视化图（含分段 R²、残差嵌套图）并保存到 CSV 同目录。用法：/icp-calibration <CSV文件路径>
argument-hint: <CSV文件路径>
---

用户提供了一个 Agilent ICP-OES 导出的 CSV 文件路径：**$ARGUMENTS**

请严格按照以下步骤执行：

---

## 第一步：读取并解析 CSV

用 Python 解析该 CSV 文件，提取所有 Type 为 `STD` 和 `BLK` 的行（即标准曲线点），重点关注以下两条谱线：

- **Ca 393.366**（Element 列含 `Ca 393`）
- **Na 588.995**（Element 列含 `Na 588.995`）

提取字段：`Label`（浓度标签）、`Concentration`（ppm）、`Intensity`（counts）。

注意：
- Blank 行对应 0 ppm
- `####` 表示仪器溢出，跳过这些点不参与拟合
- `Uncal` 表示超出校准范围，同样跳过

解析代码示例：

```python
import csv, re, sys
sys.stdout.reconfigure(encoding='utf-8')

rows = []
with open(r'$ARGUMENTS', encoding='utf-8', errors='ignore') as f:
    reader = csv.reader(f)
    for row in reader:
        if len(row) < 10:
            continue
        label, typ, _, element, _, _, _, conc, unit, intensity = row[:10]
        if typ not in ('STD', 'BLK'):
            continue
        if '####' in conc or 'Uncal' in conc or conc.strip() == '':
            continue
        try:
            rows.append({
                'label': label.strip(),
                'type': typ.strip(),
                'element': element.strip(),
                'conc': float(conc),
                'intensity': float(intensity),
            })
        except ValueError:
            continue

# 按谱线分组
ca_data = [(r['conc'], r['intensity']) for r in rows if 'Ca 393' in r['element']]
na_data = [(r['conc'], r['intensity']) for r in rows if 'Na 588.995' in r['element']]
ca_data.sort(); na_data.sort()
print('Ca points:', ca_data)
print('Na points:', na_data)
```

---

## 第二步：分段线性拟合

对 Ca 和 Na 各做**两段连续线性拟合**（breakpoint 由 scipy 自动优化）：

```python
import numpy as np
from scipy.optimize import curve_fit
from scipy.stats import linregress

def piecewise_linear(x, xb, k1, b1, k2):
    b2 = k1 * xb + b1 - k2 * xb
    return np.where(x <= xb, k1 * x + b1, k2 * x + b2)

def fit_piecewise(conc, intensity, guess_break):
    conc = np.array(conc, dtype=float)
    intensity = np.array(intensity, dtype=float)
    k0, b0, *_ = linregress(conc, intensity)
    popt, _ = curve_fit(
        piecewise_linear, conc, intensity,
        p0=[guess_break, k0, b0, k0 * 0.75],
        bounds=([conc[1], 0, -np.inf, 0], [conc[-2], np.inf, np.inf, np.inf]),
        maxfev=20000
    )
    return popt

def r2_score(y, yhat):
    y, yhat = np.array(y), np.array(yhat)
    return 1 - np.sum((y - yhat)**2) / np.sum((y - y.mean())**2)
```

Ca 断点初始猜测：20 ppm；Na 断点初始猜测：40 ppm。

---

## 第三步：生成可视化图

图像规格：
- 1 行 2 列子图，figsize=(15, 6.5)，dpi=180
- 左图：Ca 393.366 nm，右图：Na 588.995 nm
- 每张子图包含：
  - 蓝色散点：STD 测量数据
  - 红色实线：Seg 1 拟合（含方程标签）
  - 橙色虚线：Seg 2 拟合（含方程标签）
  - 绿色菱形 + 竖虚线：断点位置
  - 左中黄色文本框：`Seg 1  R² = x.xxxxxx` 和 `Seg 2  R² = x.xxxxxx`（分段分别计算）
  - 右下嵌套残差柱状图（inset_axes）

分段 R² 计算方式：
```python
mask1 = conc <= xb
mask2 = conc >= xb
r2_seg1 = r2_score(intensity[mask1], piecewise_linear(conc[mask1], *popt))
r2_seg2 = r2_score(intensity[mask2], piecewise_linear(conc[mask2], *popt))
```

R² 标注位置：`ax.text(0.04, 0.44, ...)` with `transform=ax.transAxes`

---

## 第四步：保存输出

- **图片**：保存为 `calibration_piecewise_fit.png`，路径与输入 CSV 文件**同目录**
- **Python 脚本**：保存为 `plot_calibration.py`，同目录，包含完整可复现代码（使用 `os.path.dirname(os.path.abspath(__file__))` 作为输出路径，不硬编码路径；移除 `matplotlib.use('Agg')`，保留 `plt.show()`）

路径获取方式：
```python
import os
csv_path = r'$ARGUMENTS'
out_dir = os.path.dirname(os.path.abspath(csv_path))
```

---

## 第五步：汇报结果

完成后输出：

1. Ca 和 Na 各自的断点位置（ppm）
2. 两段拟合方程（斜率 + 截距）
3. 两段各自的 R²
4. 图片和脚本保存路径

---

## 注意事项

- 使用 `matplotlib.use('Agg')` 仅在后台生成图片时使用；若生成 py 脚本供用户运行，脚本中**不要**加此行，保留 `plt.show()`
- f-string 中避免使用 `·`、`²` 等特殊字符直接打印到终端（GBK 编码问题），图中标注可用 `R$^2$` LaTeX 语法
- Ca 396.847 nm 在高浓度下普遍溢出，优先使用 Ca 393.366 nm
- Na 568.263 / 568.821 nm 通常显示 `Uncal`，优先使用 Na 588.995 nm
