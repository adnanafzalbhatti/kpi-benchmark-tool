"""Chart Engine for rendering standardized corporate KPI benchmark charts.

Provides unified styling, implementation date vertical marker with badge,
pre/post period shading, and benchmark mean reference lines.
"""

from __future__ import annotations

import io
from typing import List, Optional, Tuple

import matplotlib
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
import numpy as np

from src.data_loader import DataLoader, KPIStatistics


class ChartEngine:
    """Generates standardized, publication-quality benchmark charts."""

    # Corporate Executive Palette
    COLOR_BG = "#FFFFFF"
    COLOR_PLOT_BG = "#FAFBFC"
    COLOR_LINE = "#1B365D"          # Professional Executive Navy
    COLOR_MARKER = "#1B365D"
    COLOR_PRE_SHADE = "#EFF6FF"      # Soft Light Blue
    COLOR_POST_SHADE = "#F0FDF4"     # Soft Light Mint Green
    COLOR_IMPL_LINE = "#DC2626"      # Vivid Crimson Red
    COLOR_IMPL_BADGE_BG = "#FEF2F2"
    COLOR_IMPL_BADGE_BORDER = "#DC2626"
    COLOR_PRE_BENCH = "#64748B"      # Slate Gray
    COLOR_POST_BENCH = "#2563EB"     # Royal Blue
    COLOR_GRID = "#E2E8F0"
    COLOR_TEXT = "#1E293B"

    def __init__(self):
        matplotlib.rcParams.update({
            "font.sans-serif": ["Segoe UI", "Calibri", "Arial", "DejaVu Sans"],
            "font.family": "sans-serif",
            "axes.edgecolor": "#CBD5E1",
            "axes.linewidth": 1.0,
        })

    def create_figure(
        self,
        loader: DataLoader,
        kpi_name: str,
        pre_start: str,
        pre_end: str,
        impl_date: str,
        post_start: str,
        post_end: str,
        figsize: Tuple[float, float] = (11.0, 5.0),
        dpi: int = 150,
        show_title_in_chart: bool = False,
    ) -> Figure:
        """Creates a standalone Matplotlib Figure for the KPI benchmark (thread-safe)."""
        stats = loader.compute_kpi_statistics(kpi_name, pre_start, pre_end, post_start, post_end)
        dates, vals = loader.get_timeline_data(kpi_name)

        fig = Figure(figsize=figsize, dpi=dpi, facecolor=self.COLOR_BG)
        FigureCanvasAgg(fig)
        ax = fig.add_subplot(1, 1, 1)
        ax.set_facecolor(self.COLOR_PLOT_BG)

        # Filter out completely None values for bounds
        valid_indices = [i for i, v in enumerate(vals) if v is not None and not np.isnan(v)]
        if not valid_indices:
            ax.text(
                0.5, 0.5, f"No valid numerical data for '{kpi_name}'",
                ha="center", va="center", transform=ax.transAxes,
                fontsize=14, color="#94A3B8"
            )
            return fig

        x_indices = list(range(len(dates)))
        numeric_vals = [np.nan if v is None else v for v in vals]

        # Calculate Y-axis limits with comfortable headroom for annotations
        valid_nums = [v for v in numeric_vals if not np.isnan(v)]
        y_min = min(valid_nums)
        y_max = max(valid_nums)
        y_range = y_max - y_min
        if y_range == 0:
            y_range = max(1.0, abs(y_max) * 0.1)
        
        # 22% headroom on top and 10% on bottom
        plot_y_min = y_min - (y_range * 0.10)
        plot_y_max = y_max + (y_range * 0.26)
        ax.set_ylim(plot_y_min, plot_y_max)

        # 1. Shading for Pre and Post Periods
        pre_idx_list = [i for i, d in enumerate(dates) if pre_start <= d <= pre_end]
        post_idx_list = [i for i, d in enumerate(dates) if post_start <= d <= post_end]

        if pre_idx_list:
            pre_x0 = pre_idx_list[0] - 0.4
            pre_x1 = pre_idx_list[-1] + 0.4
            ax.axvspan(pre_x0, pre_x1, color=self.COLOR_PRE_SHADE, alpha=0.6, zorder=1)
            # Top label for Pre-Period
            pre_mid = (pre_x0 + pre_x1) / 2.0
            ax.text(
                pre_mid, plot_y_max - (y_range * 0.03),
                f"PRE-PERIOD (N={stats.pre_count})",
                ha="center", va="top", fontsize=8.5, fontweight="bold",
                color="#475569", zorder=3,
                bbox=dict(boxstyle="round,pad=0.25", facecolor="#FFFFFF", edgecolor="#CBD5E1", alpha=0.9)
            )

        if post_idx_list:
            post_x0 = post_idx_list[0] - 0.4
            post_x1 = post_idx_list[-1] + 0.4
            ax.axvspan(post_x0, post_x1, color=self.COLOR_POST_SHADE, alpha=0.6, zorder=1)
            # Top label for Post-Period
            post_mid = (post_x0 + post_x1) / 2.0
            ax.text(
                post_mid, plot_y_max - (y_range * 0.03),
                f"POST-PERIOD (N={stats.post_count})",
                ha="center", va="top", fontsize=8.5, fontweight="bold",
                color="#166534", zorder=3,
                bbox=dict(boxstyle="round,pad=0.25", facecolor="#FFFFFF", edgecolor="#86EFAC", alpha=0.9)
            )

        # 2. Benchmark Horizontal Mean Lines
        if stats.has_valid_pre and pre_idx_list:
            p_start = pre_idx_list[0] - 0.35
            p_end = pre_idx_list[-1] + 0.35
            ax.hlines(
                y=stats.pre_mean, xmin=p_start, xmax=p_end,
                colors=self.COLOR_PRE_BENCH, linestyles="--", linewidth=1.8,
                label=f"Pre Avg: {stats.pre_mean:,.2f}", zorder=4
            )
            # Label on benchmark line
            ax.text(
                p_end, stats.pre_mean, f" Pre Avg: {stats.pre_mean:,.2f}",
                va="center", ha="left", fontsize=8.5, fontweight="bold",
                color=self.COLOR_PRE_BENCH, zorder=5
            )

        if stats.has_valid_post and post_idx_list:
            p_start = post_idx_list[0] - 0.35
            p_end = post_idx_list[-1] + 0.35
            ax.hlines(
                y=stats.post_mean, xmin=p_start, xmax=p_end,
                colors=self.COLOR_POST_BENCH, linestyles="--", linewidth=1.8,
                label=f"Post Avg: {stats.post_mean:,.2f}", zorder=4
            )
            # Label on benchmark line
            ax.text(
                p_end, stats.post_mean, f" Post Avg: {stats.post_mean:,.2f}",
                va="center", ha="left", fontsize=8.5, fontweight="bold",
                color=self.COLOR_POST_BENCH, zorder=5
            )

        # 3. Main Data Curve & Points
        ax.plot(
            x_indices, numeric_vals,
            color=self.COLOR_LINE, linewidth=2.4,
            marker="o", markersize=6.5,
            markerfacecolor=self.COLOR_MARKER,
            markeredgecolor="#FFFFFF", markeredgewidth=1.5,
            label=kpi_name, zorder=6
        )

        # 4. Implementation Date Marker & Callout Badge
        impl_idx = None
        if impl_date in dates:
            impl_idx = dates.index(impl_date)
            # Vertical dashed line
            ax.axvline(
                x=impl_idx, color=self.COLOR_IMPL_LINE,
                linestyle="--", linewidth=2.2, zorder=7
            )
            # Prominent badge at top
            ax.text(
                impl_idx, plot_y_max - (y_range * 0.02),
                f"IMPLEMENTATION: {impl_date}",
                ha="center", va="top", fontsize=9.5, fontweight="bold",
                color=self.COLOR_IMPL_LINE, zorder=8,
                bbox=dict(
                    boxstyle="round,pad=0.35",
                    facecolor=self.COLOR_IMPL_BADGE_BG,
                    edgecolor=self.COLOR_IMPL_BADGE_BORDER,
                    linewidth=1.4,
                )
            )
            # Highlight data point on implementation date if available
            impl_val = numeric_vals[impl_idx]
            if not np.isnan(impl_val):
                ax.plot(
                    [impl_idx], [impl_val],
                    marker="D", markersize=9,
                    markerfacecolor=self.COLOR_IMPL_LINE,
                    markeredgecolor="#FFFFFF", markeredgewidth=2.0,
                    zorder=9
                )

        # 5. Gridlines & Axes Formatting
        ax.grid(True, linestyle="--", linewidth=0.7, color=self.COLOR_GRID, alpha=0.8, zorder=2)
        ax.set_xticks(x_indices)
        # Format date labels cleanly
        date_labels = [d.replace("-", "/") for d in dates]
        ax.set_xticklabels(date_labels, rotation=35, ha="right", fontsize=9, color=self.COLOR_TEXT)
        ax.tick_params(axis="both", which="both", length=4, color="#94A3B8")

        # Formatter for Y-axis numbers
        def y_axis_formatter(x, pos):
            if abs(x) >= 1_000_000_000:
                return f"{x*1e-9:.1f}B"
            elif abs(x) >= 1_000_000:
                return f"{x*1e-6:.1f}M"
            elif abs(x) >= 10_000:
                return f"{x*1e-3:.0f}k"
            elif abs(x) >= 100:
                return f"{x:,.0f}"
            elif abs(x) >= 1:
                return f"{x:.2f}"
            else:
                return f"{x:.3f}"

        ax.yaxis.set_major_formatter(ticker.FuncFormatter(y_axis_formatter))
        ax.tick_params(axis="y", labelsize=9, labelcolor=self.COLOR_TEXT)

        # 6. Overall Performance Delta Badge
        if show_title_in_chart:
            delta_sign = "+" if stats.abs_delta > 0 else ""
            delta_color = "#166534" if stats.abs_delta > 0 else ("#991B1B" if stats.abs_delta < 0 else "#475569")
            summary_text = (
                f"Pre Avg: {stats.pre_mean:,.2f}   |   Post Avg: {stats.post_mean:,.2f}   |   "
                f"Delta: {delta_sign}{stats.abs_delta:,.2f} ({delta_sign}{stats.pct_delta:.2f}%)"
            )
            ax.text(
                0.02, 0.05, summary_text,
                transform=ax.transAxes, ha="left", va="bottom",
                fontsize=9.5, fontweight="bold", color=delta_color,
                bbox=dict(
                    boxstyle="round,pad=0.35",
                    facecolor="#FFFFFF",
                    edgecolor="#CBD5E1",
                    linewidth=1.0,
                    alpha=0.95
                ),
                zorder=8
            )

        if show_title_in_chart:
            ax.set_title(kpi_name, fontsize=14, fontweight="bold", pad=12, color="#0F172A")

        ax.set_xlim(-0.6, len(dates) - 0.4)
        fig.tight_layout()
        return fig

    def render_to_png_bytes(
        self,
        loader: DataLoader,
        kpi_name: str,
        pre_start: str,
        pre_end: str,
        impl_date: str,
        post_start: str,
        post_end: str,
        dpi: int = 300,
    ) -> bytes:
        """Renders the chart to in-memory PNG bytes at high resolution for PPTX embedding."""
        fig = self.create_figure(
            loader=loader,
            kpi_name=kpi_name,
            pre_start=pre_start,
            pre_end=pre_end,
            impl_date=impl_date,
            post_start=post_start,
            post_end=post_end,
            figsize=(11.5, 4.8),
            dpi=dpi,
            show_title_in_chart=False,  # PPTX has slide title at top, center-aligned
        )
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor=fig.get_facecolor())
        buf.seek(0)
        return buf.getvalue()
