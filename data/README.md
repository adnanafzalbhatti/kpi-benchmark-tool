# Dataset Directory (`data/`)

Place your network performance KPI and counter data files in this directory.

### Supported File Formats:
- **Comma-Separated Values** (`.csv`)
- **Microsoft Excel** (`.xlsx`, `.xls`)

### Required Structure:
1. **Date Column**: A column containing sequential dates or timestamps (e.g., `Date`, `RecordDate`, `Timestamp`, `Period`, etc.).
2. **KPI Columns**: Numeric columns representing network quality, traffic, or counter metrics (e.g. `RRC_Setup_Success_Rate (%)`, `ERAB_Drop_Rate (%)`, `DL_User_Throughput_Mbps`, etc.).

### Included Sample:
- [`sample_network_kpis.csv`](sample_network_kpis.csv): A realistic, synthetic 21-day trial dataset demonstrating cellular network optimization pre/post performance cutover.
