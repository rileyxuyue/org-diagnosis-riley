# 部门敬满报告自动生成工具

## 目录结构

```
report_tool/
├── config.py                   # ⭐ 配置文件（修改文件路径）
├── generate_report.py          # ⭐ 主入口（运行这个）
├── data_loader.py              # 数据加载与格式化工具
├── module1_core_table.py       # 模块1：核心维度表
├── module2_risk_table.py       # 模块2：风险区间表
├── module3_chart.py            # 模块3：柱状图
├── module4_bottom_table.py     # 模块4：排名靠后明细表
├── module5_subdivision_table.py # 模块5：细分项大表
├── report_builder.js           # Word文档组装器（Node.js）
├── data/                       # 放入三个Excel数据文件
│   ├── 数据结构样本.xlsx         # 部门数据（1451列）
│   ├── BG相关数据样本.xlsx       # BG数据
│   └── 题目与标题对照表.xlsx      # 题目对照表
└── output/                     # 生成的报告保存在这里
```

---

## 快速开始

### 第一步：安装依赖

```bash
# Python 依赖
pip install pandas openpyxl matplotlib

# Node.js 依赖
npm install -g docx
```

### 第二步：放入数据文件

将三个 Excel 文件放入 `data/` 目录：
- 部门数据 Excel
- BG数据 Excel
- 题目与标题对照表 Excel

### 第三步：修改 config.py

打开 `config.py`，将文件路径改为你的实际文件名：

```python
DEPT_FILE = "data/你的部门数据文件.xlsx"
BG_FILE   = "data/你的BG数据文件.xlsx"
VAR_FILE  = "data/你的题目对照表文件.xlsx"
```

### 第四步：生成报告

```bash
# 生成单个部门报告
python3 generate_report.py "KFC/战略发展部"

# 生成所有部门报告
python3 generate_report.py --all
```

报告将保存至 `output/` 目录，文件名格式为 `报告_部门名称.docx`。

---

## 报告结构

每份报告包含以下章节：

| 章节 | 内容 |
|------|------|
| 标题页 | 部门名称、所属BG、生成日期 |
| 2.3.1 总分与定位 | 核心维度表、风险区间表、47道题柱状图、末10%明细表 |
| 2.3.2 值得关注的细分项 | 四分组对比大表（排名差异 + 增幅差异） |
| 2.3.3 敬满开放题 | 留空 |

---

## 核心逻辑说明

### 击败率（_fav_bg）
- 0~100，越高越好，100 = BG内排名第一
- 柱状图颜色：蓝色≥50、橙色20-50、红色≤20

### 部门内排名
- 该题 `_fav` 在该部门47道子题中的排名（1=最高分）

### 细分项分组条件（差值阈值 > 10）
| 分组 | 条件 |
|------|------|
| BG排名较高的题 | 部门内排名 - BG内排名 > 10 |
| 部门排名较高的题 | BG内排名 - 部门内排名 > 10 |
| 部门增幅落后BG | BG增幅 - 部门增幅 > 10% |
| 部门增幅领先BG | 部门增幅 - BG增幅 > 10% |

---

## 常见问题

**Q：部门名称怎么查？**
运行以下命令查看所有可用部门名称：
```bash
python3 -c "from data_loader import load_all; d,_,_ = load_all(); print('\n'.join(d['部门']))"
```

**Q：去年数据缺失怎么显示？**
排名变化列显示 `/`，分值变化列显示 `-`，属正常情况。

**Q：生成报告报错？**
检查 Node.js 是否已安装（`node --version`），以及 `docx` 包是否已安装（`npm list -g docx`）。
