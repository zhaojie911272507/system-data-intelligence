# 🧠 skill-data-intelligence

> A professional **Cursor Agent Skill** for system-level file operations, deep data analysis, and professional data visualization.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)](#)
[![Cursor Skill](https://img.shields.io/badge/Cursor-Agent%20Skill-purple)](#)

---

## 📦 What This Skill Does

This Cursor Agent Skill equips the AI with three tightly integrated superpowers:

| Capability | Technologies |
|---|---|
| 🔧 **System File Access** | win32com, xlwings, AppleScript, openpyxl, python-docx |
| 📊 **Deep Data Analysis** | pandas, numpy, scipy, statsmodels |
| 🎨 **Professional Visualization** | Plotly (interactive), Matplotlib (report-quality) |

**Supported file formats:** `.xlsx` `.xls` `.xlsm` `.et` `.docx` `.doc` `.wps` `.txt` `.md` `.rtz` `.csv` `.json`

---

## 🗂️ Repository Structure

```
skill-data-intelligence/
├── SKILL.md                     ← Cursor skill entry point (≤500 lines)
├── references/
│   ├── windows-api.md           ← Windows COM / PowerShell reference
│   ├── macos-api.md             ← macOS xlwings / AppleScript reference
│   ├── file-formats.md          ← Excel / WPS / Word / RTZ format guide
│   └── viz-patterns.md          ← Chart selection & styling reference
├── scripts/
│   ├── win_excel_reader.py      ← Windows COM Excel reader
│   ├── mac_excel_reader.py      ← macOS xlwings Excel reader
│   ├── wps_extractor.py         ← WPS Spreadsheet & Writer extractor
│   ├── doc_parser.py            ← Word / TXT / Markdown / RTZ parser
│   ├── deep_analyzer.py         ← Full data analysis engine
│   └── viz_engine.py            ← Plotly + Matplotlib visualization engine
├── requirements.txt
└── README.md
```

---

## 🚀 Installation

### Personal Skill (available across all projects)

```bash
mkdir -p ~/.cursor/skills/data-intelligence
cp -r . ~/.cursor/skills/data-intelligence/
```

### Project Skill (shared with team)

```bash
mkdir -p .cursor/skills/data-intelligence
cp -r . .cursor/skills/data-intelligence/
```

### Python Dependencies

```bash
# Core (all platforms)
pip install pandas numpy scipy statsmodels plotly matplotlib openpyxl python-docx

# Windows only
pip install pywin32

# macOS only
pip install xlwings
```

---

## 💡 Trigger Scenarios

The agent will **automatically apply** this skill when you say:

- "帮我分析这个 Excel 文件..."
- "读取 WPS 表格里的数据"
- "生成一份可视化报告"
- "分析这份销售数据的趋势"
- "提取 Word 文档中的表格"
- "解析这个 RTZ 项目文件"
- "做个仪表盘展示数据"

---

## 🔄 Workflow Example

```python
# The agent orchestrates all three capabilities automatically:

# Step 1: System file access
df = read_excel_mac("/path/to/sales.xlsx")   # or read_excel_via_com() on Windows

# Step 2: Deep analysis
analyzer = DeepAnalyzer(df)
report = analyzer.run_full_analysis()

# Step 3: Visualization
fig = create_dashboard(report, title="Sales Intelligence Dashboard")
export_viz(fig, "output/sales_report", formats=["html", "png"])
```

---

## 📤 Output Format

Every analysis produces:

```
outputs/
├── report_YYYYMMDD_HHMMSS.html   ← Interactive full report
├── charts/
│   ├── overview.png
│   ├── trend.png
│   └── distribution.png
├── data/
│   ├── processed_data.csv
│   └── analysis_result.json
└── summary.md                     ← Key insights in plain text
```

---

## 📚 References

- [Windows API Reference](references/windows-api.md)
- [macOS API Reference](references/macos-api.md)
- [File Formats Guide](references/file-formats.md)
- [Visualization Patterns](references/viz-patterns.md)

---

## 📄 License

MIT License — feel free to use, modify, and distribute.