---
name: icp-merge
description: 扫描父目录下的所有日期子文件夹，每个子文件夹用自己的 plot_calibration.py 标定曲线处理当天的原始数据，最终合并输出 merged_results.xlsx 到父目录。每个子文件夹需先用 /icp-calibration 生成 plot_calibration.py。用法：/icp-merge <父目录路径>
argument-hint: <父目录路径>
---

用户指定的父目录：**$ARGUMENTS**

目录结构约定：
```
父目录/
├── 20260412/
│   ├── Na, Ca_ 20260412.csv     ← 原始数据
│   └── plot_calibration.py      ← /icp-calibration 生成
├── 20260501/
│   ├── Na, Ca_ 20260501.csv
│   └── plot_calibration.py
└── merged_results.xlsx          ← 本 skill 的输出
```

请严格按照以下步骤执行：

---

## 第一步：扫描子文件夹

扫描父目录下的所有子文件夹，找出包含 `plot_calibration.py` 的文件夹（即已完成标定的测量批次）：

```python
import os, glob, sys

parent_dir = r'$ARGUMENTS'
sessions = []

for entry in sorted(os.scandir(parent_dir), key=lambda e: e.name):
    if not entry.is_dir():
        continue
    calib = os.path.join(entry.path, 'plot_calibration.py')
    if not os.path.exists(calib):
        print(f"  跳过 {entry.name}：无 plot_calibration.py（未完成标定）")
        continue
    # 检查 __main__ 保护
    with open(calib, encoding='utf-8', errors='ignore') as f:
        src = f.read()
    if "__name__ == '__main__'" not in src and '__name__ == "__main__"' not in src:
        print(f"  警告 {entry.name}：plot_calibration.py 缺少 __main__ 保护，请用 /icp-calibration 重新生成")
        continue
    # 找子文件夹中的原始数据文件
    data_files = [f for ext in ('*.xlsx', '*.csv')
                  for f in glob.glob(os.path.join(entry.path, ext))
                  if os.path.basename(f) not in
                  {'plot_calibration.py', 'merge_and_calc.py',
                   'merged_results.xlsx', 'calibration_piecewise_fit.png'}]
    if not data_files:
        print(f"  跳过 {entry.name}：无原始数据文件")
        continue
    sessions.append({'name': entry.name, 'path': entry.path,
                     'calib': calib, 'data_files': sorted(data_files)})

print(f"\n发现 {len(sessions)} 个有效测量批次：")
for s in sessions:
    print(f"  [{s['name']}]  标定: plot_calibration.py  数据: {[os.path.basename(f) for f in s['data_files']]}")
```

将结果展示给用户，确认无误后继续。

---

## 第二步：询问各批次的配置信息

对每个发现的子文件夹，**询问用户**：

1. **Date 标签**：通常直接用子文件夹名（如 `2026-04-12`），确认或修改
2. **标签重命名规则**：是否有样品标签需要重命名（如 `R-` → `10wt%-`）；若无则跳过

示例询问格式：
```
批次 [20260412]：
  - Date 标签（默认 "2026-04-12"）：
  - 标签重命名（如 R-:10wt%-，无则回车跳过）：
```

---

## 第三步：生成并运行 merge_and_calc.py

根据用户确认的信息，在**父目录**生成 `merge_and_calc.py`。

核心逻辑：每个子文件夹用自己的 `plot_calibration.py` 构造逆函数，互不干扰：

```python
"""
ICP-OES 多批次数据合并
每个子文件夹使用自己的 plot_calibration.py 标定曲线。
"""
import io, os, re, sys, importlib.util
import numpy as np
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PARENT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── 标定加载：从指定路径动态 import plot_calibration.py ──────────────────────
def load_calibration(calib_path):
    """动态加载任意路径的 plot_calibration.py，返回逆函数字典。"""
    spec = importlib.util.spec_from_file_location('plot_calibration', calib_path)
    mod  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    def _make_inverse(popt, b2):
        xb, k1, b1, k2 = popt
        i_brk = k1 * xb + b1
        def inv(intensity):
            i = float(intensity)
            return (i - b1) / k1 if i <= i_brk else (i - b2) / k2
        return inv

    return {
        'Ca 393.366': ('Ca 393.366 nm', _make_inverse(mod.popt_ca, mod.b2_ca)),
        'Na 588.995': ('Na 588.995 nm', _make_inverse(mod.popt_na, mod.b2_na)),
    }

# ── 工具函数 ─────────────────────────────────────────────────────────────────
def _to_float(val):
    if pd.isna(val): return np.nan
    s = str(val).strip()
    if s in ('####', 'Uncal', '-', '', 'nan'): return np.nan
    try: return float(s)
    except ValueError: return np.nan

def load_raw(path):
    if path.endswith('.xlsx'):
        return pd.read_excel(path, header=2)
    with open(path, 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()
    return pd.read_csv(io.StringIO(''.join(lines[2:])))

# ── 构建样品表 ────────────────────────────────────────────────────────────────
def build_sample_table(df_raw, date_label, calibrated, label_renames=None):
    df = df_raw[df_raw['Type'] == 'Sample'].copy()
    df = df[df['Element'].isin(calibrated)].copy()
    if label_renames:
        for old, new in label_renames.items():
            df['Label'] = df['Label'].apply(
                lambda x: new + x[len(old):] if isinstance(x, str) and x.startswith(old) else x)
    df['_intensity'] = df['Intensity'].apply(_to_float)
    df['_dilution']  = df['Dilution'].apply(_to_float).fillna(1.0)
    rows = []
    for (label, dt), grp in df.groupby(['Label', 'Date Time'], sort=False):
        row = {'Date': date_label, 'Label': label, 'Measurement Time': dt}
        for elem, (col_label, inv_fn) in calibrated.items():
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
    return pd.DataFrame(rows)

# ── 各批次配置（由 /icp-merge skill 根据用户确认自动填写）────────────────────
SESSIONS = [
    # {
    #   'calib'   : r'子文件夹/plot_calibration.py',
    #   'files'   : [r'子文件夹/Na, Ca_ 20260412.csv'],
    #   'date'    : '2026-04-12',
    #   'renames' : {},          # 无重命名则留空 {}
    # },
]

# ── 处理每个批次 ──────────────────────────────────────────────────────────────
df_list = []
for s in SESSIONS:
    calibrated = load_calibration(s['calib'])
    for fpath in s['files']:
        df_raw = load_raw(fpath)
        df = build_sample_table(df_raw, s['date'], calibrated, s.get('renames'))
        df_list.append(df)
    print(f"  [{s['date']}] {sum(len(build_sample_table(load_raw(f), s['date'], load_calibration(s['calib']))) for f in s['files'])} samples")

df_combined = pd.concat(df_list, ignore_index=True)

# ── Excel 样式 ────────────────────────────────────────────────────────────────
_PALETTE = ['BDD7EE','C6EFCE','FFE699','FCE4D6','E2CCFF',
            'F4CCCC','B7DEE8','D9E1F2','FFF2CC','D6DCE4']
_HDR_FILL  = PatternFill('solid', fgColor='1F497D')
_HDR_FONT  = Font(color='FFFFFF', bold=True)
_BORDER    = Border(**{s: Side(style='thin', color='BBBBBB')
                       for s in ('left','right','top','bottom')})
_ALIGN_CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
_ALIGN_MID = Alignment(horizontal='center', vertical='center')

def _group_of(label):
    g = re.sub(r'[\s-]+\d.*$', '', str(label).strip())
    return g or label

def _assign_colors(labels):
    group_idx = {}
    for lbl in labels:
        g = _group_of(lbl)
        if g not in group_idx:
            group_idx[g] = len(group_idx) % len(_PALETTE)
    fills = {lbl: PatternFill('solid', fgColor=_PALETTE[group_idx[_group_of(lbl)]])
             for lbl in labels}
    return fills, group_idx

def style_and_write(ws, df):
    labels = df['Label'].tolist()
    fills, group_idx = _assign_colors(labels)
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

# ── 写出 ──────────────────────────────────────────────────────────────────────
out_path = os.path.join(PARENT_DIR, 'merged_results.xlsx')
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    df_combined.to_excel(writer, sheet_name='Combined', index=False)
    style_and_write(writer.sheets['Combined'], df_combined)

print(f"\nSaved -> {out_path}  ({len(df_combined)} samples total)")
```

将 `SESSIONS` 列表按用户确认的配置填写后，在父目录运行。

---

## 第四步：汇报结果

完成后输出：
1. 各批次样品数及使用的标定文件
2. 合并总样品数和分组情况
3. 输出路径（`merged_results.xlsx` 在父目录）

---

## 目录结构

```
Performance/
├── 20260404/  ← 特殊批次（共用0412标定，已处理）
├── 20260412/
│   ├── Na, Ca_ 20260412.csv
│   └── plot_calibration.py    ← /icp-calibration 生成
├── 20260501/  ← 未来新批次
│   ├── Na, Ca_ 20260501.csv
│   └── plot_calibration.py
└── merged_results.xlsx        ← 仅在用户明确要求时由 /icp-merge 生成
```

每次新测量只需运行：
```
/icp-calibration <子文件夹/当天csv>
```
**不需要立即 merge**。只有用户明确提出合并需求时，才运行 `/icp-merge <父目录>`。

---

## 注意事项

- 每个子文件夹的 `plot_calibration.py` 通过 `importlib` **动态加载**，互不干扰，各用各的标定参数
- Agilent xlsx：真实表头在第 3 行（header=2）
- Agilent csv：前 2 行路径注释，第 3 行才是表头
- 溢出（`####`）和未标定（`Uncal`）均视为 NaN
- Dilution 为空时默认为 1
