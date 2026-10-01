"""Main entry point for KPI Benchmark & Trial Analysis Tool.

Launches the Tkinter GUI application by default, or runs in CLI mode if flags are provided.
"""

import argparse
import multiprocessing
import os
import sys
from datetime import datetime

multiprocessing.freeze_support()

# Ensure project root is in sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_loader import DataLoader
from src.chart_engine import ChartEngine
from src.ppt_generator import PowerPointGenerator


def run_cli(args):
    """Executes benchmark report generation via command-line arguments."""
    file_path = args.file
    if not file_path:
        # Default to first file in data/
        data_dir = os.path.abspath("data")
        for ext in ["*.csv", "*.xlsx", "*.xls"]:
            import glob
            files = glob.glob(os.path.join(data_dir, ext))
            if files:
                file_path = files[0]
                break

    if not file_path or not os.path.exists(file_path):
        print(f"Error: Data file not found: {file_path}")
        sys.exit(1)

    print(f"[1/4] Loading data from {file_path}...")
    loader = DataLoader(file_path)
    print(f"      Detected date column: {loader.date_col}")
    print(f"      Dates available ({len(loader.available_dates)}): {loader.available_dates[0]} to {loader.available_dates[-1]}")
    print(f"      Numeric KPIs/counters discovered: {len(loader.numeric_kpis)}")

    # Implementation date and periods
    impl_date = args.impl or loader.available_dates[len(loader.available_dates) // 2]
    sug = loader.get_suggested_periods(impl_date)

    pre_start = args.pre_start or sug["pre_start"]
    pre_end = args.pre_end or sug["pre_end"]
    post_start = args.post_start or sug["post_start"]
    post_end = args.post_end or sug["post_end"]

    print(f"[2/4] Configuration:")
    print(f"      Implementation Date : {impl_date}")
    print(f"      Pre-Period          : {pre_start} to {pre_end}")
    print(f"      Post-Period         : {post_start} to {post_end}")

    # KPIs selection
    if args.kpis:
        if args.kpis.strip().lower() == "all":
            selected_kpis = loader.numeric_kpis
        else:
            selected_kpis = [k.strip() for k in args.kpis.split(",") if k.strip() in loader.numeric_kpis]
    else:
        # Pick default representative KPIs
        sample_candidates = ["RACH_Success_Rate", "DL Volume (Gbit)", "pmRaAttCbra", "Data Availability (%)", "Num of Cells"]
        selected_kpis = [k for k in sample_candidates if k in loader.numeric_kpis]
        if not selected_kpis:
            selected_kpis = loader.numeric_kpis[:5]

    print(f"[3/4] Selected {len(selected_kpis)} KPIs for report: {selected_kpis}")

    output_path = args.output
    if not output_path:
        out_dir = os.path.abspath("output")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(out_dir, f"KPI_Benchmark_Report_{ts}.pptx")
    elif os.path.isdir(output_path):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_path, f"KPI_Benchmark_Report_{ts}.pptx")

    print(f"[4/4] Generating PowerPoint deck...")
    ppt_gen = PowerPointGenerator()

    def progress(curr, total, name):
        print(f"      Slide {curr+1}/{total+1}: {name}")

    out_file = ppt_gen.generate_report(
        loader=loader,
        selected_kpis=selected_kpis,
        pre_start=pre_start,
        pre_end=pre_end,
        impl_date=impl_date,
        post_start=post_start,
        post_end=post_end,
        output_path=output_path,
        template_path=getattr(args, "template", None),
        include_summary_table=getattr(args, "summary_table", True),
        progress_callback=progress
    )
    print(f"\nSUCCESS: Report successfully created at: {out_file}")


def main():
    """Main entry point."""
    multiprocessing.freeze_support()

    # Log startup for diagnostic visibility
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for d in [os.getcwd(), os.path.dirname(os.path.abspath(sys.executable)), os.path.dirname(os.path.abspath(sys.argv[0]))]:
        try:
            with open(os.path.join(d, "kpi_tool_debug.log"), "a", encoding="utf-8") as f:
                f.write(f"[{ts}] main started, argv={sys.argv}\n")
                f.flush()
        except Exception:
            pass

    # Safeguard standard streams for frozen windowed executables
    if sys.stdout is None:
        try:
            sys.stdout = open(os.devnull, "w", encoding="utf-8")
        except Exception:
            pass
    if sys.stderr is None:
        try:
            sys.stderr = open(os.devnull, "w", encoding="utf-8")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="KPI Benchmark & Trial Analysis Tool")
    parser.add_argument("--cli", action="store_true", help="Run in command-line mode instead of launching GUI")
    parser.add_argument("--file", type=str, help="Path to input Excel or CSV data file")
    parser.add_argument("--impl", type=str, help="Implementation date (YYYY-MM-DD)")
    parser.add_argument("--pre-start", type=str, help="Pre-period start date (YYYY-MM-DD)")
    parser.add_argument("--pre-end", type=str, help="Pre-period end date (YYYY-MM-DD)")
    parser.add_argument("--post-start", type=str, help="Post-period start date (YYYY-MM-DD)")
    parser.add_argument("--post-end", type=str, help="Post-period end date (YYYY-MM-DD)")
    parser.add_argument("--kpis", type=str, help="Comma-separated KPI names or 'all'")
    parser.add_argument("--output", type=str, help="Output .pptx path or folder")
    parser.add_argument("--template", type=str, help="Path to PowerPoint template (.pptx or .potx)")
    parser.add_argument("--summary-table", dest="summary_table", action="store_true", default=True, help="Include editable summary table on Slide 1 (default: True)")
    parser.add_argument("--no-summary-table", dest="summary_table", action="store_false", help="Do not include summary table on Slide 1")

    # If --cli is given or any specific CLI flags are provided without running interactively
    if len(sys.argv) > 1 and sys.argv[1] != "":
        args = parser.parse_args()
        if args.cli or args.file or args.kpis:
            try:
                import pyi_splash
                pyi_splash.close()
            except Exception:
                pass
            run_cli(args)
            return

    # Default to GUI with safe error reporting
    try:
        from src.gui import launch_gui
        launch_gui()
    except Exception as e:
        import traceback
        err_msg = traceback.format_exc()
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "KPI Benchmark Tool - Launch Error",
                f"An unexpected error occurred while launching the application:\n\n{err_msg}"
            )
            root.destroy()
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()

