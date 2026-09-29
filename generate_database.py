"""
generate_database.py — Backward-compatible wrapper for reinspection_suite.py
===========================================================================
Delegates execution to reinspection_suite.py to generate Re-Inspection-Database-Month-Year.xlsx
or launch the Python Visualizer GUI Dashboard.
"""

import sys
import reinspection_suite

if __name__ == '__main__':
    # If called with --gui or no arguments, launch the GUI dashboard
    if '--gui' in sys.argv or len(sys.argv) == 1:
        reinspection_suite.launch_dashboard()
    else:
        import argparse
        parser = argparse.ArgumentParser(description="Generate Professional Re-Inspection Database from 3 input folders.")
        parser.add_argument('--re-dir', default='Re-Inspection', help='Path to Re-Inspection folder')
        parser.add_argument('--ftt-dir', default='FTT', help='Path to FTT folder')
        parser.add_argument('--hfpa-dir', default='HFPA', help='Path to HFPA folder')
        parser.add_argument('--output', default=None, help='Output Excel file path (defaults to Re-Inspection-Database-Month-Year.xlsx)')
        parser.add_argument('--month', default=None, help='Month label override, e.g. "June 2026"')
        parser.add_argument('--gui', action='store_true', help='Launch Graphical User Interface')
        args = parser.parse_args()

        if args.gui:
            reinspection_suite.launch_dashboard(args.re_dir, args.ftt_dir, args.hfpa_dir)
        else:
            reinspection_suite.generate_database_workbook(args.re_dir, args.ftt_dir, args.hfpa_dir, args.output, month_label=args.month)
