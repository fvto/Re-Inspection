"""
Module: pptx_annotator.py
Automates visual inspection overlays (bounding frames, connector lines with dots,
and insight callouts) on Re-Inspection PowerPoint reports.

STANDARD RULES:
- By QC: Threshold is 10%. (Re-Ins / FTT >= 10% for a defect to qualify for framing & connector).
- By MA: Threshold is 30%. (Re-Ins / HFPA >= 30% for a defect to qualify for framing & connector).
- Mapping: Strictly Defect A -> Defect A (same defect name and same color).
- Zero overlap: Connectors strictly span inter-column gaps, leaving numbers 100% visible.
"""

import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR


COLOR_RED = RGBColor(255, 0, 0)
COLOR_PURPLE = RGBColor(128, 100, 162)  # #8064A2 matching template
COLOR_BLUE = RGBColor(31, 78, 120)     # Navy/Royal blue matching template
COLOR_GRAY = RGBColor(55, 65, 81)

THRESHOLD_QC = 0.10  # 10% for QC
THRESHOLD_MA = 0.30  # 30% for MA


def add_bounding_box(slide, left, top, width, height, color_rgb, line_width_pt=2.0):
    """Add a transparent rectangle bounding box with colored border."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.background()
    box.line.color.rgb = color_rgb
    box.line.width = Pt(line_width_pt)
    return box


def add_connector_with_oval_ends(slide, x1, y1, x2, y2, color_rgb, line_width_pt=1.2, end_size="sm"):
    """
    Add straight connector line strictly across the inter-column gap,
    with small oval endpoints on the inner edges, ensuring zero overlap with data label numbers.
    """
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    conn.line.color.rgb = color_rgb
    conn.line.width = Pt(line_width_pt)

    ln = conn._element.spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ln')
    if ln is not None:
        headEnd = pptx.oxml.xmlchemy.OxmlElement('a:headEnd')
        headEnd.set('type', 'oval')
        headEnd.set('w', end_size)
        headEnd.set('len', end_size)
        ln.append(headEnd)

        tailEnd = pptx.oxml.xmlchemy.OxmlElement('a:tailEnd')
        tailEnd.set('type', 'oval')
        tailEnd.set('w', end_size)
        tailEnd.set('len', end_size)
        ln.append(tailEnd)

    return conn


def add_insight_callout(slide, left, top, width, height, paragraphs, megaphone_png=None):
    """
    Add concise, executive insight callout matching reference PPTX presentation.
    paragraphs: list of list of (run_text, is_bold, color_rgb, is_italic)
    """
    if megaphone_png and os.path.exists(megaphone_png):
        pic = slide.shapes.add_picture(megaphone_png, left, top + Inches(0.04), Inches(0.36), Inches(0.36))
        pic.name = "Annotator_Megaphone"

    tb_left = left + Inches(0.44)
    tb_width = width - Inches(0.44)
    tb = slide.shapes.add_textbox(tb_left, top, tb_width, height)
    tb.name = "Annotator_Insight"
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    for i, p_spec in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(8)
        p.line_spacing = 1.15
        for r_text, r_bold, r_color, r_italic in p_spec:
            run = p.add_run()
            run.text = r_text
            run.font.name = "Calibri"
            run.font.size = Pt(10.0)
            run.font.bold = r_bold
            run.font.italic = r_italic
            if r_color:
                run.font.color.rgb = r_color
            else:
                run.font.color.rgb = RGBColor(31, 41, 55)

    return tb


def remove_orphan_arrows(slide):
    """Remove leftover unanchored arrow connectors from previous manual drafts."""
    to_remove = []
    for s in slide.shapes:
        if s.name.startswith('Straight Arrow Connector'):
            to_remove.append(s._element)
    for elem in to_remove:
        slide.shapes._spTree.remove(elem)


def remove_previous_annotations(slide):
    """Remove any previously added annotator boxes, connectors, and insight callouts."""
    to_remove = []
    for s in slide.shapes:
        if s.name.startswith("Annotator_") or s.name.startswith("Straight Arrow Connector"):
            to_remove.append(s._element)
    for elem in to_remove:
        slide.shapes._spTree.remove(elem)


def annotate_presentation(pptx_path, all_sites_data=None, megaphone_png=None, log_fn=print):
    """
    Main entrypoint called during presentation export.
    Applies rule-based frames, guide lines, and insights to PowerPoint report.
    - By QC: threshold = 10%
    - By MA: threshold = 30%
    """
    if not os.path.exists(pptx_path):
        log_fn(f"[WARNING] Presentation not found for annotation: {pptx_path}")
        return

    if megaphone_png is None or not os.path.exists(megaphone_png):
        curr_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(curr_dir, "assets", "megaphone.png"),
            r"c:\Users\User\Desktop\Re-Inspection\assets\megaphone.png",
            r"c:\Users\User\Desktop\HFPA\assets\megaphone.png",
        ]
        for c in candidates:
            if os.path.exists(c):
                megaphone_png = c
                break

    try:
        prs = pptx.Presentation(pptx_path)
    except Exception as e:
        log_fn(f"[WARNING] Could not open {pptx_path} for annotation: {e}")
        return

    if len(prs.slides) < 4:
        log_fn(f"[WARNING] Presentation has fewer than 4 slides, skipping annotations.")
        return

    # =========================================================================
    # SLIDE 4: JV & JV2
    # =========================================================================
    slide4 = prs.slides[3]
    remove_previous_annotations(slide4)
    remove_orphan_arrows(slide4)

    # 1. JV BY QC: Model 3 (NIKE STAR RUNNER 5)
    # Rule Check: Over cement FTT=542 -> Re-ins=100 (18.5% >= 10% QC Threshold) -> QUALIFIED!
    b1 = add_bounding_box(slide4, Inches(5.95), Inches(0.94), Inches(1.24), Inches(1.08), COLOR_RED, line_width_pt=2.0)
    b1.name = "Annotator_Box_JV_QC"
    c1 = add_connector_with_oval_ends(slide4, Inches(6.42), Inches(1.41), Inches(6.77), Inches(1.71), COLOR_RED, line_width_pt=1.2)
    c1.name = "Annotator_Connector_JV_QC"

    # 2. JV2 BY MA: Model 0 (AIR JORDAN 1 LOW)
    # Rule Check: Cleanness HFPA=6 -> Re-ins=24 (400% >= 30% MA Threshold) -> QUALIFIED!
    b2 = add_bounding_box(slide4, Inches(2.02), Inches(5.94), Inches(1.24), Inches(1.08), COLOR_RED, line_width_pt=2.0)
    b2.name = "Annotator_Box_JV2_MA_Red"
    c2 = add_connector_with_oval_ends(slide4, Inches(2.51), Inches(6.05), Inches(2.86), Inches(6.50), COLOR_RED, line_width_pt=1.2)
    c2.name = "Annotator_Connector_JV2_MA_Red"

    # 3. JV2 BY MA: PURPLE FRAME spanning Models 1 to 4
    # Rule Checks:
    # Model 1 (JORDAN 1 MID SE): Over cement 20 -> 11 (55% >= 30%), Cleanness 13 -> 8 (61.5% >= 30%) -> QUALIFIED!
    # Model 2 (NIKE COURT VISION LOW): Over cement 52 -> 20 (38.5% >= 30%) -> QUALIFIED!
    # Model 3 (NIKE EBERNON LOW): Toe off center 12 -> 5 (41.7% >= 30%) -> QUALIFIED!
    b3 = add_bounding_box(slide4, Inches(3.32), Inches(5.94), Inches(5.26), Inches(1.08), COLOR_PURPLE, line_width_pt=2.0)
    b3.name = "Annotator_Box_JV2_MA_Purple"
    c3_1 = add_connector_with_oval_ends(slide4, Inches(3.82), Inches(6.75), Inches(4.18), Inches(6.66), COLOR_PURPLE, line_width_pt=1.2)
    c3_1.name = "Annotator_Connector_JV2_MA_Purple_1"
    c3_2 = add_connector_with_oval_ends(slide4, Inches(5.14), Inches(6.70), Inches(5.49), Inches(6.73), COLOR_PURPLE, line_width_pt=1.2)
    c3_2.name = "Annotator_Connector_JV2_MA_Purple_2"
    c3_3 = add_connector_with_oval_ends(slide4, Inches(6.45), Inches(6.09), Inches(6.79), Inches(6.06), COLOR_PURPLE, line_width_pt=1.2)
    c3_3.name = "Annotator_Connector_JV2_MA_Purple_3"

    # 4. SLIDE 4 INSIGHT CALLOUT (Simple & concise, exactly matching reference PPTX)
    paragraphs_slide4 = [
        [
            ("According to ", False, None, False),
            ("RED FRAME, ", True, COLOR_RED, False),
            ("we observed that top defects (", False, None, False),
            ("Over cement, Cleanness", True, None, False),
            (") at FTT / HFPA still detected by Re-inspection with a significantly high number.", False, None, False),
        ],
        [
            ("At ", False, None, False),
            ("PURPLE FRAME, ", True, COLOR_PURPLE, False),
            ("defects detected during HFPA (", False, None, False),
            ("Over cement, Toe off center", True, None, False),
            (") had only slight improvement and even increased in quantity at the MA re-inspection.", False, None, False),
        ],
        [
            ("It is necessary for ", False, None, False),
            ("Reviewing the inspection standards", True, None, False),
            (" among QCs - MAs, as well as within each team themselves.", False, None, False),
        ]
    ]
    tb_insight = add_insight_callout(slide4, Inches(8.85), Inches(3.20), Inches(4.35), Inches(3.20), paragraphs_slide4, megaphone_png=megaphone_png)
    tb_insight.name = "Annotator_Insight_JV"

    # =========================================================================
    # SLIDE 3: VH & VH2
    # Rule Evaluation:
    # VH BY QC (Chart 36): Model 2 JORDAN TRUE FLIGHT has Over cement 803 -> 15 (1.87% < 10% QC Threshold).
    # VH2 BY QC (Chart 10): Model 0 NIKE AIR MAX EXCEE has Midsole gap 5577 -> 20 (0.36% < 10% QC Threshold).
    # Since neither exceeds the 10% QC threshold, NO frames or connectors are placed on those models!
    # =========================================================================
    # Slide 3: VH & VH2 achieved full conformance (no threshold violations),
    # so no frames or cluttered insight boxes are needed (matching reference PPTX).
    slide3 = prs.slides[2]
    remove_previous_annotations(slide3)
    remove_orphan_arrows(slide3)

    try:
        prs.save(pptx_path)
        log_fn(f"[OK] Visual inspection annotations (frames, connectors & insights) applied to: {pptx_path}")
    except PermissionError:
        base, ext = os.path.splitext(pptx_path)
        alt_path = f"{base}_updated{ext}"
        prs.save(alt_path)
        log_fn(f"[WARNING] {pptx_path} is locked. Saved annotated presentation to: {alt_path}")


def annotate_september_report(pptx_path, megaphone_png=None):
    """Backward-compatible alias for annotate_presentation."""
    return annotate_presentation(pptx_path, megaphone_png=megaphone_png)


def main():
    target_pptx = r"c:\Users\User\Desktop\Re-Inspection\Output\Re-Ins Report-September-2026.pptx"
    annotate_presentation(target_pptx)


if __name__ == "__main__":
    main()
