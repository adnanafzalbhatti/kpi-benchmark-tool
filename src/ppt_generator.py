"""PowerPoint Report Generator for KPI Benchmark & Trial Analysis.

Produces a standardized 16:9 widescreen presentation with:
- Executive Title Slide
- 1 Slide per KPI with center-aligned KPI title at the top
- Consistent corporate styling and layout
- High-resolution embedded benchmark charts with implementation markers
"""

from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Callable, List, Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt

from src.chart_engine import ChartEngine
from src.data_loader import DataLoader, KPIStatistics


class PowerPointGenerator:
    """Generates corporate PowerPoint deck for KPI benchmark trials."""

    # Corporate Colors
    COLOR_PRIMARY = RGBColor(27, 54, 93)      # Corporate Navy #1B365D
    COLOR_SECONDARY = RGBColor(37, 99, 235)   # Royal Blue #2563EB
    COLOR_DARK_TEXT = RGBColor(30, 41, 59)    # Slate #1E293B
    COLOR_MUTED_TEXT = RGBColor(100, 116, 139)# Gray #64748B
    COLOR_BG_CARD = RGBColor(241, 245, 249)   # Light Slate #F1F5F9
    COLOR_BORDER = RGBColor(203, 213, 225)    # Border #CBD5E1
    COLOR_ACCENT_RED = RGBColor(220, 38, 38)  # Implementation Red #DC2626
    COLOR_ACCENT_GREEN = RGBColor(22, 101, 52)# Success Green #166534

    # Summary Table Colors (matching corporate benchmark table)
    COLOR_TABLE_HEADER_BLUE = RGBColor(207, 226, 243)  # Soft Light Blue #CFE2F3
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_BLACK = RGBColor(0, 0, 0)

    def __init__(self, chart_engine: Optional[ChartEngine] = None):
        self.chart_engine = chart_engine or ChartEngine()

    def _find_clean_layout(self, prs: Presentation):
        """Finds a suitable blank/content layout in the presentation."""
        for layout in prs.slide_layouts:
            if "blank" in layout.name.lower():
                return layout
        for layout in prs.slide_layouts:
            if len(layout.placeholders) == 0:
                return layout
        if len(prs.slide_layouts) > 6:
            return prs.slide_layouts[6]
        return prs.slide_layouts[-1]

    def generate_report(
        self,
        loader: DataLoader,
        selected_kpis: List[str],
        pre_start: str,
        pre_end: str,
        impl_date: str,
        post_start: str,
        post_end: str,
        output_path: str,
        template_path: Optional[str] = None,
        include_summary_table: bool = True,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
    ) -> str:
        """Builds and saves the complete PowerPoint benchmark report."""
        if not selected_kpis:
            raise ValueError("No KPIs selected for report generation.")

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

        using_template = False
        if template_path and os.path.exists(template_path):
            try:
                prs = Presentation(template_path)
                using_template = True
            except PermissionError:
                raise PermissionError(
                    f"Template '{os.path.basename(template_path)}' is currently open in PowerPoint.\n\n"
                    "Please save and close the template in PowerPoint, or clear the Template field to generate without a template."
                )
        else:
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

        clean_layout = self._find_clean_layout(prs)

        # 1. Slide 1: Summary Table (if requested) or Title Slide
        summary_slide_count = 0
        if include_summary_table:
            summary_slide_count = self._build_summary_table_slide(
                prs=prs,
                blank_layout=clean_layout,
                loader=loader,
                selected_kpis=selected_kpis,
                pre_start=pre_start,
                pre_end=pre_end,
                impl_date=impl_date,
                post_start=post_start,
                post_end=post_end,
                using_template=using_template,
            )
        else:
            self._build_title_slide(
                prs=prs,
                blank_layout=clean_layout,
                loader=loader,
                selected_kpis=selected_kpis,
                pre_start=pre_start,
                pre_end=pre_end,
                impl_date=impl_date,
                post_start=post_start,
                post_end=post_end,
                using_template=using_template,
            )
            summary_slide_count = 1

        total_kpis = len(selected_kpis)
        total_slides = summary_slide_count + total_kpis

        # 2. KPI Slides (1 slide per KPI)
        for idx, kpi_name in enumerate(selected_kpis, 1):
            if progress_callback:
                progress_callback(idx, total_kpis, kpi_name)

            self._build_kpi_slide(
                prs=prs,
                blank_layout=clean_layout,
                loader=loader,
                kpi_name=kpi_name,
                pre_start=pre_start,
                pre_end=pre_end,
                impl_date=impl_date,
                post_start=post_start,
                post_end=post_end,
                slide_num=summary_slide_count + idx,
                total_slides=total_slides,
                using_template=using_template,
            )

        prs.save(output_path)
        return output_path

    @staticmethod
    def _set_cell_border(cell, color: str = "000000", width: str = "12700") -> None:
        """Applies solid borders to all 4 sides of a table cell."""
        tcPr = cell._tc.get_or_add_tcPr()
        for border_name in ["lnL", "lnR", "lnT", "lnB"]:
            border_xml = parse_xml(
                f'<a:{border_name} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width}" cmpd="s">\n'
                f'  <a:solidFill>\n'
                f'    <a:srgbClr val="{color}"/>\n'
                f'  </a:solidFill>\n'
                f'</a:{border_name}>'
            )
            tcPr.append(border_xml)

    def _build_summary_table_slide(
        self,
        prs: Presentation,
        blank_layout,
        loader: DataLoader,
        selected_kpis: List[str],
        pre_start: str,
        pre_end: str,
        impl_date: str,
        post_start: str,
        post_end: str,
        using_template: bool = False,
    ) -> int:
        """Constructs native editable PowerPoint table slide(s) on Slide 1 for all checked KPIs."""
        chunk_size = 16
        kpi_chunks = [selected_kpis[i:i + chunk_size] for i in range(0, len(selected_kpis), chunk_size)]
        total_chunks = len(kpi_chunks) if kpi_chunks else 1
        if not kpi_chunks:
            kpi_chunks = [[]]

        for chunk_idx, kpis_in_chunk in enumerate(kpi_chunks, 1):
            slide = prs.slides.add_slide(blank_layout)

            # Top Accent Header Bar (only for default blank layout)
            if not using_template:
                accent_bar = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(0), Inches(0), Inches(13.333), Inches(0.35)
                )
                accent_bar.fill.solid()
                accent_bar.fill.fore_color.rgb = self.COLOR_PRIMARY
                accent_bar.line.fill.background()

            # Slide Title & Subtitle Header
            title_box = slide.shapes.add_textbox(
                Inches(0.8), Inches(0.45), Inches(11.733), Inches(0.95)
            )
            tf = title_box.text_frame
            tf.word_wrap = True
            p1 = tf.paragraphs[0]
            part_suffix = f" (Part {chunk_idx}/{total_chunks})" if total_chunks > 1 else ""
            p1.text = f"KPI Benchmark Summary Table{part_suffix}"
            p1.font.name = "Calibri"
            p1.font.size = Pt(22)
            p1.font.bold = True
            p1.font.color.rgb = self.COLOR_PRIMARY
            p1.alignment = PP_ALIGN.CENTER

            p2 = tf.add_paragraph()
            p2.text = f"Pre-Period: {pre_start} to {pre_end}   |   Implementation: {impl_date}   |   Post-Period: {post_start} to {post_end}"
            p2.font.name = "Calibri"
            p2.font.size = Pt(11)
            p2.font.color.rgb = self.COLOR_MUTED_TEXT
            p2.alignment = PP_ALIGN.CENTER

            # Table Dimensions
            rows = len(kpis_in_chunk) + 1
            cols = 4
            left = Inches(1.266)
            top = Inches(1.50)
            width = Inches(10.8)
            available_h = 5.5
            calc_h = min(available_h, max(1.8, 0.32 * rows))

            table_shape = slide.shapes.add_table(rows, cols, left, top, width, Inches(calc_h))
            table = table_shape.table

            table.columns[0].width = Inches(5.4)
            table.columns[1].width = Inches(1.8)
            table.columns[2].width = Inches(1.8)
            table.columns[3].width = Inches(1.8)

            # Header Row
            headers = ["KPIs", "Pre", "Post", "Change%"]
            for col_idx, h_text in enumerate(headers):
                cell = table.cell(0, col_idx)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Inches(0.08)
                cell.margin_right = Inches(0.08)
                cell.margin_top = Inches(0.03)
                cell.fill.solid()
                if col_idx == 0:
                    cell.fill.fore_color.rgb = self.COLOR_TABLE_HEADER_BLUE
                else:
                    cell.fill.fore_color.rgb = self.COLOR_WHITE

                self._set_cell_border(cell, color="000000", width="12700")

                p = cell.text_frame.paragraphs[0]
                p.text = h_text
                p.font.name = "Calibri"
                p.font.size = Pt(11)
                p.font.bold = True
                p.font.color.rgb = self.COLOR_BLACK
                p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT

            # Data Rows
            for r_idx, kpi_name in enumerate(kpis_in_chunk, 1):
                try:
                    stats = loader.compute_kpi_statistics(
                        kpi_name=kpi_name,
                        pre_start=pre_start,
                        pre_end=pre_end,
                        post_start=post_start,
                        post_end=post_end,
                    )
                    pre_val = stats.pre_mean
                    post_val = stats.post_mean

                    name_lower = kpi_name.lower()
                    is_rate = any(kw in name_lower for kw in [
                        "%", "rate", "ratio", "succ", "sr", "loss", "bler", "accessibility", "retainability", "dtx"
                    ]) or (0 <= pre_val <= 100 and 0 <= post_val <= 100 and any(kw in name_lower for kw in ["pct", "prb", "cce", "util", "avail"]))

                    if is_rate:
                        change_val = post_val - pre_val
                    else:
                        if pre_val != 0:
                            change_val = ((post_val - pre_val) / abs(pre_val)) * 100.0
                        else:
                            change_val = 0.0

                    if abs(pre_val) >= 1000 and abs(pre_val - round(pre_val, 2)) < 1e-4 and round(pre_val, 2) == round(pre_val, 0):
                        pre_str = f"{int(round(pre_val))}"
                    else:
                        pre_str = f"{pre_val:.2f}"

                    if abs(post_val) >= 1000 and abs(post_val - round(post_val, 2)) < 1e-4 and round(post_val, 2) == round(post_val, 0):
                        post_str = f"{int(round(post_val))}"
                    else:
                        post_str = f"{post_val:.2f}"

                    change_str = f"{change_val:.2f}"
                except Exception:
                    pre_str, post_str, change_str = "N/A", "N/A", "N/A"

                row_vals = [kpi_name, pre_str, post_str, change_str]
                for c_idx, val_str in enumerate(row_vals):
                    cell = table.cell(r_idx, c_idx)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    cell.margin_left = Inches(0.08)
                    cell.margin_right = Inches(0.08)
                    cell.margin_top = Inches(0.03)
                    cell.margin_bottom = Inches(0.03)

                    cell.fill.solid()
                    if c_idx == 0:
                        cell.fill.fore_color.rgb = self.COLOR_TABLE_HEADER_BLUE
                    else:
                        cell.fill.fore_color.rgb = self.COLOR_WHITE

                    self._set_cell_border(cell, color="000000", width="12700")

                    p = cell.text_frame.paragraphs[0]
                    p.text = val_str
                    p.font.name = "Calibri"
                    p.font.size = Pt(10) if len(kpis_in_chunk) > 12 else Pt(10.5)
                    p.font.bold = (c_idx == 0)
                    p.font.color.rgb = self.COLOR_BLACK
                    p.alignment = PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.RIGHT

        return total_chunks

    def _build_title_slide(
        self,
        prs: Presentation,
        blank_layout,
        loader: DataLoader,
        selected_kpis: List[str],
        pre_start: str,
        pre_end: str,
        impl_date: str,
        post_start: str,
        post_end: str,
        using_template: bool = False,
    ) -> None:
        """Constructs the executive title and metadata slide."""
        slide = prs.slides.add_slide(blank_layout)

        # Top Accent Header Bar (only for default blank layout)
        if not using_template:
            accent_bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0), Inches(0), Inches(13.333), Inches(0.35)
            )
            accent_bar.fill.solid()
            accent_bar.fill.fore_color.rgb = self.COLOR_PRIMARY
            accent_bar.line.fill.background()

        # Title Box
        title_box = slide.shapes.add_textbox(
            Inches(1.0), Inches(1.1), Inches(11.333), Inches(1.4)
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = "KPI Benchmark & Trial Analysis Report"
        p1.font.name = "Calibri"
        p1.font.size = Pt(36)
        p1.font.bold = True
        p1.font.color.rgb = self.COLOR_PRIMARY
        p1.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = "Pre vs. Post Implementation Performance Evaluation"
        p2.font.name = "Calibri"
        p2.font.size = Pt(18)
        p2.font.color.rgb = self.COLOR_MUTED_TEXT
        p2.alignment = PP_ALIGN.CENTER

        # Metadata Card Background
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(1.5), Inches(2.9), Inches(10.333), Inches(3.4)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = self.COLOR_BG_CARD
        card.line.color.rgb = self.COLOR_BORDER
        card.line.width = Pt(1.5)

        # Metadata Text inside Card
        meta_box = slide.shapes.add_textbox(
            Inches(1.8), Inches(3.1), Inches(9.733), Inches(3.0)
        )
        mtf = meta_box.text_frame
        mtf.word_wrap = True

        data_source_name = os.path.basename(loader.file_path)
        items = [
            ("Source Dataset", data_source_name),
            ("Pre-Period Window", f"{pre_start}  to  {pre_end}"),
            ("Implementation Date", f"{impl_date} (Trial Cutover)"),
            ("Post-Period Window", f"{post_start}  to  {post_end}"),
            ("Evaluated KPIs / Counters", f"{len(selected_kpis)} metrics selected"),
            ("Report Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ]

        for i, (label, val) in enumerate(items):
            p = mtf.paragraphs[0] if i == 0 else mtf.add_paragraph()
            p.space_after = Pt(8)
            p.alignment = PP_ALIGN.LEFT
            
            run_lbl = p.add_run()
            run_lbl.text = f"{label:<28} :   "
            run_lbl.font.bold = True
            run_lbl.font.size = Pt(14)
            run_lbl.font.color.rgb = self.COLOR_PRIMARY

            run_val = p.add_run()
            run_val.text = val
            run_val.font.bold = (label == "Implementation Date")
            run_val.font.size = Pt(14)
            if label == "Implementation Date":
                run_val.font.color.rgb = self.COLOR_ACCENT_RED
            else:
                run_val.font.color.rgb = self.COLOR_DARK_TEXT

    def _build_kpi_slide(
        self,
        prs: Presentation,
        blank_layout,
        loader: DataLoader,
        kpi_name: str,
        pre_start: str,
        pre_end: str,
        impl_date: str,
        post_start: str,
        post_end: str,
        slide_num: int,
        total_slides: int,
        using_template: bool = False,
    ) -> None:
        """Constructs an individual KPI benchmark slide."""
        slide = prs.slides.add_slide(blank_layout)
        stats = loader.compute_kpi_statistics(kpi_name, pre_start, pre_end, post_start, post_end)

        # 1. Slide Title: KPI Name, CENTER-ALIGNED at the top
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.65)
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = kpi_name
        p.alignment = PP_ALIGN.CENTER  # Explicit user requirement: Center-aligned
        p.font.name = "Calibri"
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_PRIMARY

        # 2. Executive Metric Ribbon (Card directly below title)
        ribbon = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(1.5), Inches(1.05), Inches(10.333), Inches(0.45)
        )
        ribbon.fill.solid()
        ribbon.fill.fore_color.rgb = self.COLOR_BG_CARD
        ribbon.line.color.rgb = self.COLOR_BORDER
        ribbon.line.width = Pt(1.0)

        ribbon_tf = ribbon.text_frame
        ribbon_tf.word_wrap = False
        ribbon_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        rp = ribbon_tf.paragraphs[0]
        rp.alignment = PP_ALIGN.CENTER

        delta_prefix = "+" if stats.abs_delta > 0 else ""
        delta_str = f"{delta_prefix}{stats.abs_delta:,.2f} ({delta_prefix}{stats.pct_delta:.2f}%)"

        metrics_line = (
            f"Pre Avg: {stats.pre_mean:,.2f} (N={stats.pre_count})   |   "
            f"Post Avg: {stats.post_mean:,.2f} (N={stats.post_count})   |   "
            f"Delta: {delta_str}   |   "
            f"Impl Date: {impl_date}"
        )
        r_run = rp.add_run()
        r_run.text = metrics_line
        r_run.font.name = "Calibri"
        r_run.font.size = Pt(11.5)
        r_run.font.bold = True
        r_run.font.color.rgb = self.COLOR_DARK_TEXT

        # 3. High-Resolution Embedded Benchmark Chart
        chart_png_bytes = self.chart_engine.render_to_png_bytes(
            loader=loader,
            kpi_name=kpi_name,
            pre_start=pre_start,
            pre_end=pre_end,
            impl_date=impl_date,
            post_start=post_start,
            post_end=post_end,
            dpi=300,
        )

        chart_stream = io.BytesIO(chart_png_bytes)
        slide.shapes.add_picture(
            chart_stream,
            left=Inches(0.8),
            top=Inches(1.58),
            width=Inches(11.733),
            height=Inches(5.25),
        )

        # 4. Slide Footer (only add default footer if not using custom company template)
        if not using_template:
            footer_box = slide.shapes.add_textbox(
                Inches(0.8), Inches(7.0), Inches(11.733), Inches(0.4)
            )
            ftf = footer_box.text_frame
            fp = ftf.paragraphs[0]
            fp.text = f"KPI Benchmark Trial Report  |  Implementation: {impl_date}  |  Slide {slide_num} of {total_slides}"
            fp.font.name = "Calibri"
            fp.font.size = Pt(9.5)
            fp.font.color.rgb = self.COLOR_MUTED_TEXT
            fp.alignment = PP_ALIGN.RIGHT
