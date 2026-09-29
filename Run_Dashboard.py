"""
Run_Dashboard.py - Re-Inspection & Quality Intelligence Suite Visualizer
======================================================================
Interactive Native Desktop Application for Quality Audit Analytics,
Multi-Station Cross Analysis, and One-Click Corporate Report Generation.
"""

import os
import sys
import argparse

# Ensure local directory is in Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    import reinspection_suite
except ImportError as e:
    print(f"[ERROR] Failed to import reinspection_suite: {e}")
    print("Please make sure reinspection_suite.py is in the same directory.")
    sys.exit(1)

def main():
    parser = argparse.ArgumentParser(
        description="Launch Re-Inspection & Quality Intelligence Suite Desktop Visualizer"
    )
    parser.add_argument('--re-dir', default='Re-Inspection', help='Path to Re-Inspection data directory')
    parser.add_argument('--ftt-dir', default='FTT', help='Path to FTT data directory')
    parser.add_argument('--hfpa-dir', default='HFPA', help='Path to HFPA data directory')
    parser.add_argument('--template-pptx', default='RE-INS REPORT.APR.2026.pptx', help='Template PowerPoint presentation path')
    parser.add_argument('--color-template', default='Color_Template.xlsx', help='Color Template Excel file path')

    args = parser.parse_args()

    # Verify input directories or fallback to detected folders
    re_dir = os.path.join(current_dir, args.re_dir) if not os.path.isabs(args.re_dir) else args.re_dir
    ftt_dir = os.path.join(current_dir, args.ftt_dir) if not os.path.isabs(args.ftt_dir) else args.ftt_dir
    hfpa_dir = os.path.join(current_dir, args.hfpa_dir) if not os.path.isabs(args.hfpa_dir) else args.hfpa_dir
    tpl_pptx = os.path.join(current_dir, args.template_pptx) if not os.path.isabs(args.template_pptx) else args.template_pptx
    color_tpl = os.path.join(current_dir, args.color_template) if not os.path.isabs(args.color_template) else args.color_template

    print("=" * 70)
    print("  🚀 Starting Re-Inspection & Quality Intelligence Suite Dashboard...")
    print("=" * 70)
    print(f"  • Workspace:      {current_dir}")
    print(f"  • Re-Inspection:  {re_dir}")
    print(f"  • FTT Folder:     {ftt_dir}")
    print(f"  • HFPA Folder:    {hfpa_dir}")
    print(f"  • PPTX Template:  {tpl_pptx}")
    print(f"  • Color Template: {color_tpl}")
    print("=" * 70)
    print("  Launching interactive GUI window...")

    reinspection_suite.launch_dashboard(
        default_re=re_dir,
        default_ftt=ftt_dir,
        default_hfpa=hfpa_dir,
        default_template=tpl_pptx,
        default_color_tpl=color_tpl
    )

if __name__ == '__main__':
    main()
