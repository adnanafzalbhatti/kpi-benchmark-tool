# KPI Benchmark & Trial Analysis Tool

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg)]()
[![Code Style](https://img.shields.io/badge/code%20style-PEP8-orange.svg)]()

> **Automated Pre/Post Performance Evaluation & Executive Slide Deck Generation for Technical Trials, Network Changes, and Engineering Audits.**

---

## 📌 Overview

When conducting technical field trials, software upgrades, or network optimization experiments, engineering teams often spend hours manually slicing data in Excel and copy-pasting charts into PowerPoint. 

The **KPI Benchmark & Trial Analysis Tool** automates this entire end-to-end workflow:
1. Ingests raw time-series performance data (`.csv`, `.xlsx`, `.xls`).
2. Automatically identifies the time-series date column and discovers numeric KPIs & counters.
3. Automatically slices data into **Pre-Trial** and **Post-Trial** comparison windows around an **Implementation Date**.
4. Computes statistical benchmarks (Mean, Median, Standard Deviation, Absolute Delta, and Percentage Change).
5. Generates an executive-ready, widescreen **PowerPoint deck (`.pptx`)**:
   - **Slide 1**: A **100% native, fully editable executive summary table** matching corporate audit styling.
   - **Slides 2+**: One dedicated widescreen slide per selected metric with high-resolution benchmark charts, shaded trial windows, cutover markers, and statistical summary ribbons.
   - **Enterprise Template Support**: Automatically inherits corporate slide masters, color schemes, logos, and typography when a company `.pptx` or `.potx` template is supplied.

---

## ✨ Key Features

- **📊 100% Native, Editable PowerPoint Tables**:
  Unlike tools that paste static image screenshots of tables, this application generates native PowerPoint table shapes (`slide.shapes.add_table`). Every cell, value, font size, and background color remains fully editable directly in Microsoft PowerPoint.
- **📈 Professional Presentation Charts**:
  - Implementation date dashed marker with date label.
  - Shaded comparison zones (Soft Blue for Pre-period, Soft Mint for Post-period).
  - Horizontal benchmark reference lines showing Pre-mean vs. Post-mean.
  - Top-centered slide titles and standardized executive palettes.
- **📄 Smart Table Pagination**:
  Selecting large batches of KPIs (e.g. 20, 50, or 100+) automatically splits the executive summary table across multiple clean slides (up to 16 rows per slide) to ensure crisp legibility without text crowding.
- **🏢 Enterprise Template Inheritance**:
  Place your company presentation template (`.pptx` or `.potx`) in the `templates/` folder. The tool preserves the slide master layout, corporate logos, headers, and legal footers.
- **🖥️ Intuitive Dual-Pane Desktop GUI**:
  Built with Tkinter and Matplotlib. Features live search filtering across hundreds of metrics, "Select All / Clear All" helpers, date picker selectors, and an embedded interactive chart preview pane.
- **⚡ Command-Line Automation (CLI)**:
  Run headless in batch scripts, scheduled jobs, or CI/CD pipelines using dedicated CLI flags.
- **🚀 Portable & Standalone Executable**:
  Packageable via PyInstaller into an instant 1-second launcher or a standalone `.exe` requiring zero Python installation on end-user machines.

---

## 🏗️ Architecture

```
kpi_benchmarks_trials/
├── main.py                     # Unified entry point (GUI launcher & CLI orchestrator)
├── src/
│   ├── data_loader.py          # Data ingestion, date normalization, & statistical computations
│   ├── chart_engine.py         # Matplotlib rendering, styling, & shaded window generation
│   ├── ppt_generator.py        # OpenXML PowerPoint slide assembly & table formatting
│   └── gui.py                  # Responsive Tkinter desktop interface with embedded preview
├── data/
│   ├── sample_network_kpis.csv # Synthetic 21-day demonstration trial dataset
│   └── README.md
├── templates/
│   └── README.md               # Instructions for custom corporate slide master templates
├── tests/
│   ├── test_summary_table.py   # Test suite for native table generation & pagination
│   └── test_pptx.py            # Presentation alignment and OpenXML structure validation
├── assets/
│   └── splash.png              # Application splash screen branding
├── requirements.txt            # Python dependencies
├── LICENSE                     # MIT Open Source License
└── Launch_KPI_Tool.bat         # 1-click launcher for Windows
```

---

## 🚀 Quick Start

### 1. Prerequisites
- Python 3.10 or higher
- Git

### 2. Installation
```bash
# Clone the repository
git clone https://github.com/<your-username>/kpi-benchmark-tool.git
cd kpi-benchmark-tool

# Create and activate a virtual environment
python -m venv .venv

# Windows (Command Prompt or PowerShell):
.venv\Scripts\activate

# Linux / macOS:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 3. Launch Desktop GUI
```bash
python main.py
```
*Alternatively, on Windows double-click `Launch_KPI_Tool.bat`.*

---

## 💻 Command-Line (CLI) Usage

Automate report generation without launching the graphical interface:

```bash
python main.py --cli \
  --file "data/sample_network_kpis.csv" \
  --kpis "RRC_Setup_Success_Rate (%),ERAB_Drop_Rate (%),DL_User_Throughput_Mbps" \
  --impl "2026-09-10" \
  --pre-start "2026-09-01" \
  --pre-end "2026-09-09" \
  --post-start "2026-09-11" \
  --post-end "2026-09-21" \
  --summary-table \
  --output "output/Trial_Benchmark_Report.pptx"
```

### CLI Arguments:
| Argument | Description | Default |
|---|---|---|
| `--cli` | Runs in headless command-line mode | `False` |
| `--file` | Path to CSV or Excel file | First file found in `data/` |
| `--kpis` | Comma-separated list of KPI column names | First 10 discovered KPIs |
| `--impl` | Implementation cutover date (`YYYY-MM-DD`) | Midpoint of dataset dates |
| `--pre-start` | Pre-trial evaluation start date | Inferred from dataset |
| `--pre-end` | Pre-trial evaluation end date | Day prior to implementation |
| `--post-start` | Post-trial evaluation start date | Day after implementation |
| `--post-end` | Post-trial evaluation end date | Inferred from dataset |
| `--template` | Optional path to custom `.pptx` template | `None` (uses clean executive layout) |
| `--summary-table` / `--no-summary-table` | Toggle Slide 1 editable summary table | `--summary-table` (Enabled) |
| `--output` | Destination path for the `.pptx` deck | `output/KPI_Benchmark_Report_<timestamp>.pptx` |

---

## 🧪 Running Automated Tests

Run the built-in test suite to verify data loading, OpenXML schema compliance, and PowerPoint table pagination:

```bash
python tests/test_summary_table.py
python tests/test_pptx.py
```

---

## 📦 Building Standalone Executables (Windows)

To package the tool into a standalone desktop application for colleagues without Python:

```bash
# Build fast, pre-extracted application directory
pyinstaller KPI_Benchmark_App.spec --noconfirm

# Or build a single portable .exe file
pyinstaller KPI_Benchmark_Tool.spec --noconfirm
```

---

## 📄 License

This project is licensed under the [MIT License](LICENSE) - see the LICENSE file for details.
