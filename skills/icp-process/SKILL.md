---
name: icp-process
description: Agilent ICP-OES 全流程一键处理：读取原始导出文件 → 提取STD拟合标准曲线（保存PNG和plot_calibration.py）→ 用标定曲线计算样品浓度 → 输出分组着色的 results.xlsx。用法：/icp-process <原始文件路径（csv或xlsx）>
argument-hint: <原始文件路径>
---

用户提供的原始 ICP-OES 导出文件：**$ARGUMENTS**

请严格按照以下步骤执行，**全程用 Python 一次性完成**，无需中断询问（除非文件读取失败）。

---

## 第一步：读取原始文件

根据文件扩展名选择读取方式：

```python
import io, os, re, sys, importlib.util
import numpy as np
import pandas as pd
from scipy.optimize import curve_fit
from scipy.stats import linregress
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

raw_path = r'$ARGUMENTS'
ext = os.path.splitext(raw_path)[1].lower()
out_dir = os.path.dirname(os.path.abspath(raw_path))

if ext == '.xlsx':
    df_raw = pd.read_excel(raw_path, header=2)        # Agilent xlsx：表头在第3行
elif ext == '.csv':
    with open(raw_path, 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()
    df_raw = pd.read_csv(io.StringIO(''.join(lines[2:])))  # 跳过前2行路径注释
else:
    print(f"不支持的文件格式：{ext}")
    sys.exit(1)

print(f"读取完成：{os.path.basename(raw_path)}，共 {len(df_raw)} 行")
print(f"包含类型：{df_raw['Type'].unique().tolist()}")
```

---

## 第二步：提取STD/BLK点，拟合标准曲线

**只使用两条有标定意义的谱线：Ca 393.366 nm 和 Na 588.995 nm**

其他谱线（Ca 396.847、Na 568.263/821/589.592、Mg 系列）忽略。

```python
def to_float(val):
    if pd.isna(val): return np.nan
    s = str(val).strip()
    if s in ('####', 'Uncal', '-', '', 'nan'): return np.nan
    try: return float(s)
    except ValueError: return np.nan

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

# 提取STD和BLK行
std_df = df_raw[df_raw['Type'].isin(['STD', 'BLK'])].copy()
std_df['_conc']  = std_df['Concentration'].apply(to_float)
std_df['_inten'] = std_df['Intensity'].apply(to_float)
std_df = std_df.dropna(subset=['_conc', '_inten'])

def get_std_points(element_pattern):
    rows = std_df[std_df['Element'].str.contains(element_pattern, na=False)]
    pts = sorted(zip(rows['_conc'].tolist(), rows['_inten'].tolist()))
    return [p[0] for p in pts], [p[1] for p in pts]

ca_conc, ca_inten = get_std_points('Ca 393')
na_conc, na_inten = get_std_points('Na 588.995')

print(f"Ca STD 点数：{len(ca_conc)}，Na STD 点数：{len(na_conc)}")

# 拟合
popt_ca = fit_piecewise(ca_conc, ca_inten, 20.0)
popt_na = fit_piecewise(na_conc, na_inten, 40.0)

xb_ca, k1_ca, b1_ca, k2_ca = popt_ca
b2_ca = k1_ca * xb_ca + b1_ca - k2_ca * xb_ca

xb_na, k1_na, b1_na, k2_na = popt_na
b2_na = k1_na * xb_na + b1_na - k2_na * xb_na

print(f"Ca 断点：{xb_ca:.2f} ppm，Na 断点：{xb_na:.2f} ppm")
```

---

## 第三步：保存标定曲线图（calibration_piecewise_fit.png）

图像规格与 `/icp-calibration` 完全一致：1行2列，figsize=(15, 6.5)，dpi=180，含分段R²标注和残差嵌套图。

```python
def plot_calibration(ax, conc, inten, popt, title):
    conc  = np.array(conc)
    inten = np.array(inten)
    xb, k1, b1, k2 = popt
    b2 = k1 * xb + b1 - k2 * xb
    yb = k1 * xb + b1
    x1 = np.linspace(conc[0], xb, 300)
    x2 = np.linspace(xb, conc[-1], 300)

    ax.scatter(conc, inten, color='#2166ac', s=70, zorder=6,
               edgecolors='white', linewidths=0.9, label='STD data')
    ax.plot(x1, k1*x1+b1, color='#d73027', lw=2.2,
            label=f'Seg1 (<={xb:.1f}ppm)\ny={k1:.3e}x+{b1:.3e}')
    ax.plot(x2, k2*x2+b2, color='#f46d43', lw=2.2, linestyle='--',
            label=f'Seg2 (>{xb:.1f}ppm)\ny={k2:.3e}x+{b2:.3e}')
    ax.axvline(xb, color='#1a9641', ls=':', lw=1.5, alpha=0.9)
    ax.scatter([xb], [yb], color='#1a9641', s=110, zorder=7,
               marker='D', label=f'Breakpoint: {xb:.1f}ppm')

    mask1, mask2 = conc <= xb, conc >= xb
    r2_1 = r2_score(inten[mask1], piecewise_linear(conc[mask1], *popt))
    r2_2 = r2_score(inten[mask2], piecewise_linear(conc[mask2], *popt))
    ax.text(0.04, 0.44, f'Seg1 R$^2$={r2_1:.6f}\nSeg2 R$^2$={r2_2:.6f}',
            transform=ax.transAxes, va='top', fontsize=9.5,
            bbox=dict(boxstyle='round,pad=0.3', fc='lightyellow', alpha=0.9))

    ax_in = ax.inset_axes([0.59, 0.04, 0.39, 0.32])
    resid = inten - piecewise_linear(conc, *popt)
    ax_in.bar(conc, resid, width=max(conc)*0.045, color='#2166ac', alpha=0.65)
    ax_in.axhline(0, color='k', lw=0.8)
    ax_in.set_xlabel('Conc (ppm)', fontsize=7)
    ax_in.set_ylabel('Residual', fontsize=7)
    ax_in.tick_params(labelsize=6)
    ax_in.set_title('Residuals', fontsize=7.5, pad=2)
    ax_in.yaxis.get_major_formatter().set_scientific(True)
    ax_in.yaxis.get_major_formatter().set_powerlimits((0,0))

    ax.set_xlabel('Concentration (ppm)', fontsize=12)
    ax.set_ylabel('Intensity (counts)', fontsize=12)
    ax.set_title(title, fontsize=13, fontweight='bold')
    ax.legend(fontsize=8.5, loc='upper left', framealpha=0.85)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.yaxis.get_major_formatter().set_scientific(True)
    ax.yaxis.get_major_formatter().set_powerlimits((0,0))

# 从文件名提取日期（如 20260412）用于图标题
date_str = re.search(r'2026\d{4}', os.path.basename(raw_path))
date_str = date_str.group() if date_str else os.path.basename(raw_path)

fig, axes = plt.subplots(1, 2, figsize=(15, 6.5))
fig.suptitle(f'ICP-OES Standard Calibration Curves – Piecewise Linear Fit  ({date_str})',
             fontsize=13, fontweight='bold', y=1.02)
plot_calibration(axes[0], ca_conc, ca_inten, popt_ca, 'Ca 393.366 nm')
plot_calibration(axes[1], na_conc, na_inten, popt_na, 'Na 588.995 nm')
plt.tight_layout()
png_path = os.path.join(out_dir, 'calibration_piecewise_fit.png')
plt.savefig(png_path, dpi=180, bbox_inches='tight')
plt.close()
print(f"标定图已保存：{png_path}")
```

---

## 第四步：保存 plot_calibration.py

将拟合参数写入 `plot_calibration.py`，含 `if __name__ == '__main__':` 保护，供 `/icp-merge` 将来使用（如需跨批次合并）。

文件结构参照现有 `plot_calibration.py` 格式，将第二步拟合出的数值硬编码进去（不再重新拟合），确保 import 无副作用。

---

## 第五步：计算样品浓度

用拟合逆函数将 Intensity → 浓度（ppm），乘稀释倍数得实际浓度：

```python
def make_inverse(popt, b2):
    xb, k1, b1, k2 = popt
    i_brk = k1 * xb + b1
    def inv(i):
        i = float(i)
        return (i - b1) / k1 if i <= i_brk else (i - b2) / k2
    return inv

inv_ca = make_inverse(popt_ca, b2_ca)
inv_na = make_inverse(popt_na, b2_na)

CALIBRATED = {
    'Ca 393.366': ('Ca 393.366 nm', inv_ca),
    'Na 588.995': ('Na 588.995 nm', inv_na),
}

df_samp = df_raw[df_raw['Type'] == 'Sample'].copy()
df_samp = df_samp[df_samp['Element'].isin(CALIBRATED)].copy()
df_samp['_intensity'] = df_samp['Intensity'].apply(to_float)
df_samp['_dilution']  = df_samp['Dilution'].apply(to_float).fillna(1.0)

# 从文件名推断日期标签
date_label = re.search(r'(\d{4}-\d{2}-\d{2}|\d{8})', os.path.basename(raw_path))
date_label = date_label.group() if date_label else date_str

rows = []
for (label, dt), grp in df_samp.groupby(['Label', 'Date Time'], sort=False):
    row = {'Date': date_label, 'Label': label, 'Measurement Time': dt}
    for elem, (col_label, inv_fn) in CALIBRATED.items():
        sub   = grp[grp['Element'] == elem]
        inten = sub['_intensity'].values[0] if not sub.empty else np.nan
        dil   = sub['_dilution'].values[0]  if not sub.empty else 1.0
        row[f'{col_label} Intensity']            = inten
        row[f'{col_label} Dilution']              = dil
        if np.isnan(inten):
            row[f'{col_label} Conc_ppm (undiluted)'] = np.nan
            row[f'{col_label} Conc_ppm (actual)']    = np.nan
        else:
            c = inv_fn(inten)
            row[f'{col_label} Conc_ppm (undiluted)'] = round(c, 4)
            row[f'{col_label} Conc_ppm (actual)']    = round(c * dil, 4)
    rows.append(row)

df_result = pd.DataFrame(rows)
print(f"样品数：{len(df_result)}")
```

---

## 第六步：输出 results.xlsx（分组着色）

```python
_PALETTE = ['BDD7EE','C6EFCE','FFE699','FCE4D6','E2CCFF',
            'F4CCCC','B7DEE8','D9E1F2','FFF2CC','D6DCE4']
_HDR_FILL  = PatternFill('solid', fgColor='1F497D')
_HDR_FONT  = Font(color='FFFFFF', bold=True)
_BORDER    = Border(**{s: Side(style='thin', color='BBBBBB')
                       for s in ('left','right','top','bottom')})
_ALIGN_CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
_ALIGN_MID = Alignment(horizontal='center', vertical='center')

def group_of(label):
    g = re.sub(r'[\s-]+\d.*$', '', str(label).strip())
    return g or label

def assign_colors(labels):
    group_idx = {}
    for lbl in labels:
        g = group_of(lbl)
        if g not in group_idx:
            group_idx[g] = len(group_idx) % len(_PALETTE)
    fills = {lbl: PatternFill('solid', fgColor=_PALETTE[group_idx[group_of(lbl)]])
             for lbl in labels}
    return fills, group_idx

def style_sheet(ws, df):
    labels = df['Label'].tolist()
    fills, group_idx = assign_colors(labels)
    for cell in ws[1]:
        cell.fill, cell.font = _HDR_FILL, _HDR_FONT
        cell.alignment, cell.border = _ALIGN_CTR, _BORDER
    for row_idx, label in enumerate(labels, start=2):
        for cell in ws[row_idx]:
            cell.fill, cell.border, cell.alignment = fills[label], _BORDER, _ALIGN_MID
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 28)
    ws.freeze_panes = 'A2'
    legend_row = ws.max_row + 2
    ws.cell(row=legend_row, column=1, value='Group Legend').font = Font(bold=True)
    for i, (group, idx) in enumerate(group_idx.items(), start=1):
        cell = ws.cell(row=legend_row + i, column=1, value=group)
        cell.fill = PatternFill('solid', fgColor=_PALETTE[idx])
        cell.border, cell.alignment, cell.font = _BORDER, _ALIGN_MID, Font(bold=True)

out_xlsx = os.path.join(out_dir, 'results.xlsx')
with pd.ExcelWriter(out_xlsx, engine='openpyxl') as writer:
    df_result.to_excel(writer, sheet_name='Results', index=False)
    style_sheet(writer.sheets['Results'], df_result)

print(f"结果已保存：{out_xlsx}")
```

---

## 第七步：汇报结果

完成后输出：
1. 标定曲线断点（Ca / Na）和分段 R²
2. 样品总数及各组分布
3. 生成的文件列表：
   - `calibration_piecewise_fit.png`
   - `plot_calibration.py`
   - `results.xlsx`

---

## 三个 Skill 的分工

| Skill | 用途 | 触发时机 |
|---|---|---|
| `/icp-process <文件>` | **常规全流程**：标定 + 计算浓度 + 输出表格 | 每次新测量后 |
| `/icp-calibration <文件>` | 单独查看/重做标定曲线 | 需要检查曲线质量时 |
| `/icp-merge <父目录>` | 跨批次合并 | 明确需要汇总多次测量时 |

---

## 注意事项

- Agilent xlsx：表头在第 3 行（header=2）
- Agilent csv：前 2 行为路径注释，跳过
- 溢出（`####`）和未标定（`Uncal`）均视为 NaN
- Dilution 为空时默认为 1
- 标签重命名（如 `R-` → `10wt%-`）若有需要，运行前告知
