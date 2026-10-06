# Re-Inspection & Quality Intelligence Suite — Evolution Ledger

## 1. Evolution Ledger Overview
This ledger tracks the chronological evolution, schema migrations, algorithm modifications, and structural changes of the **Re-Inspection & Quality Intelligence Suite**.
When changes occur in upstream standalone code, templates, or business rules, the companion extraction prompt `MCP_TOOL_DELTA_EXTRACTOR.md` uses this document and Git commit history (`git diff v1.0.0..HEAD`) to perform 3-way AST/code delta audits and produce incremental migration plans.

---

## 2. Baseline Record — Version 1.0.0

- **Baseline Tag**: `v1.0.0`
- **Release Name**: `Release v1.0.0: Initial extracted baseline for MCP Server integration`
- **Extraction Date**: `2026-09-29`
- **Repository Branch**: `master`
- **Operating Environment**: Windows 10/11 x64, Python 3.10-3.13 (x64)
- **Baseline Commit SHA**: `55008a12ec59fa44c68438d3405ca971132c5b63`
- **Baseline Commit Reference**: Staged via Conventional Commits (`feat(core)`, `feat(templates)`, `feat(agents)`, `docs(integrate)`) and tagged at `v1.0.0`.

### 2.1 Cryptographic Fingerprint Manifest (Baseline)

| File Path | Role | Size (Bytes) | SHA-256 Checksum |
|---|---|---|---|
| `reinspection_suite.py` | Core Analytical Engine & Excel Builders | 95,086 | `cee499edcdf11de821c5e4e083144ef0801fff5977ef4061281277a6470caf26` |
| `pptx_generator.py` | OpenXML Presentation Exporter | 50,997 | `e0a208eafb4e847b1292b9458497f5377470c47aaac2dffe02ae0d69baca19d6` |
| `Color_Template.xlsx` | Defect Color Palette Master (148 colors) | 12,177 | `076389448b23df6309e63b790fd885ddca0e153b24b6b137645f792a28746b84` |
| `RE-INS REPORT.APR.2026.pptx` | Executive Presentation Master Template | 649,767 | `e4922a0be8799f39e3c7cde90446b22dccf9b7bd53b7f5f83109fdd0f31d88f8` |
| `Combined.py` | Backward-Compatibility CLI Wrapper | 1,437 | `df5bb8dd43e9440e615a4e0da787cb1832e2be6e22d181365d3487bb6604c0ab` |
| `generate_database.py` | Backward-Compatibility CLI Wrapper | 1,599 | `dd6aab59bb36b9956a099a4f7e348ed9ff051aa9d403222ea60f0ac72096c2cb` |
| `Generate_Database.ts` | Office Scripts (M365 Database Engine) | 52,777 | `b9b67e97403c0165db67451823081c53aee4e17eb73f038a16c28becf1bd50dc` |
| `RecycleReport.ts` | Office Scripts (M365 Recycle Engine) | 19,740 | `ba4f2dffd9689fb22d285305d4f975aa3463614360950849758bca99c99c679b` |
| `Combined_RecycleReport.ts` | Office Scripts (M365 Combined Engine) | 22,948 | `c6173b25667bb4c2752363676e38fd7add8529b2b4d10c9228e28a964e6b775c` |
| `requirements.txt` | Package Dependencies | 235 | `0a5559bd0b88d56300ca048847af5d8bc43c1ba9b834f15c197b05557bf50514` |
| `Run_Dashboard.py` | Tkinter Dashboard Launcher | 2,853 | `14fd5a619521cfedf4f932a06b2c638854d5fe10fe9a5edf7ff2372d330c64d1` |
| `install_dependencies.py` | Automated Dependency Provisioner | 9,422 | `acb1cc34da94844eb2af84251df64bfaa5db3ad60bfab62d36dc2d4a0a32cd61` |

### 2.2 Core Contracts Baseline Snapshot

1. **Input Directory Contracts**:
   - `Re-Inspection/`: Monthly `.xls`/`.xlsx` audit workbooks matching `*VH*.xls*`, `*VH2*.xls*`, `*JV*.xls*`, `*JV2*.xls*`.
     - Marker Column: `Remark` (all columns to the right are defect categories).
     - Filtering: `Category == 'QC'` (Initial Audit) and `Category == 'MA'` (Re-Audit).
   - `FTT/`: Assembly defect logs (`.xlsx`).
     - Key fields: `Model` (or `Model 2`), `Defect Issues`, `Issues Q'ty`, `Station`.
   - `HFPA/`: High frequency audit logs (`.xlsx`).
     - Marker Column: `Defect Q'ty (Total Defect Q'ty)`. Filter: `Pass/Fail == 'Fail'`.

2. **Generated Deliverable Contracts**:
   - `Re-Inspection-Database-[Period].xlsx`: 7 worksheets (`RE-VH`, `RE-VH2`, `RE-JV`, `RE-JV2`, `VN-INDO`, `QC`, `MA`). Dynamic Excel formulas and styling.
   - `Recycle_Report-[Period].xlsx`: 4 factory worksheets (`VH`, `VH2`, `JV`, `JV2`), each structured into 7 sections with Defect Share percentages (`TL%`).
   - `Re-Ins Report-[Period].pptx`: 5 widescreen executive slides containing 19 OpenXML DrawingML charts with synchronized defect hex colors.

3. **Defect Palette Contract**:
   - Master palette: `Color_Template.xlsx` containing 148 registered defects.
   - Dynamic algorithm: Golden ratio hue stepping + Euclidean distance check ($\ge 36.0$) to assign non-clashing ARGB colors for newly encountered defect categories.

---

## 4. Version 1.1.0 Delta Record — Visual Automation & Dynamic Titles

- **Version Tag**: `v1.1.0`
- **Release Name**: `Release v1.1.0: Add dynamic presentation titles, date cover badge, pptx annotator, assets and latest monthly datasets`
- **Extraction Date**: `2026-10-06`
- **Repository Branch**: `main`
- **Operating Environment**: Windows 10/11 x64, Python 3.10-3.13 (x64)
- **Delta Commit SHA**: `b0eea0e`
- **Previous Baseline Tag**: `v1.0.0` (`55008a1`)

### 4.1 Delta Modifications & Enhancements

1. **Dynamic Presentation Header & Cover Date**:
   - `pptx_generator.py`: Introduced `_update_slide_header_title` to dynamically update presentation slide titles across all slides based on `title_format` (default: `Re-Inspection Report_{month_label}`).
   - `pptx_generator.py`: Introduced `_update_cover_slide` to format and inject dynamic date badges on Slide 1 cover (`Date: <Month> <Year>`).
   - `reinspection_suite.py`: Extended `generate_presentation_report` and `generate_all_reports` to accept and propagate `title_format`.

2. **Automated Visual Inspection Annotator (`pptx_annotator.py`)**:
   - Automated generation of bounding frames, connector lines with oval terminal points, and insight callouts with megaphone icon.
   - Enforced standard quality thresholds:
     - **QC Threshold**: $\ge 10\%$ ratio of Re-Inspection vs. FTT defect count.
     - **MA Threshold**: $\ge 30\%$ ratio of Re-Inspection vs. HFPA defect count.
     - **Defect Mapping Rule**: Strict one-to-one mapping (`Defect A -> Defect A`) preserving color identity.
     - **Zero Overlap Layout**: Connectors span inter-column gaps without obscuring numeric data labels.

3. **Graphic Assets & Palette Updates**:
   - Added `assets/megaphone.png` for executive presentation insight boxes.
   - Updated `Color_Template.xlsx` with recent monthly defect color definitions.

### 4.2 Cryptographic Fingerprint Manifest (Delta v1.1.0)

| File Path | Role | Size (Bytes) | SHA-256 Checksum |
|---|---|---|---|
| `pptx_generator.py` | OpenXML Presentation Exporter (Dynamic Titles) | 53,136 | `446933dfd920e543b5e4070a96f1ec2d01ea313010b9576c62c2f4fc911b3337` |
| `reinspection_suite.py` | Core Analytical Engine & Pipeline | 95,242 | `3464ac86ca982fa18151f15aa879a97d7ee4d84f2b0bbff99596b6fb5434d708` |
| `pptx_annotator.py` | Visual Inspection Overlays & Insights Engine | 10,931 | `8727a897a13a7c6b579fb647f2ef8c72877a561dc1ba96739678147d341b52a4` |
| `Color_Template.xlsx` | Defect Color Palette Master (Updated) | 12,296 | `601a4d25bebde8c83a1ea4489bf16c7cb5f448eaad1d6540d9990833919cbbd8` |
| `assets/megaphone.png` | Icon Asset for PowerPoint Callout Shapes | 11,252 | `4c50deee82deae811b7d5ee8210334ccae3053805ae6522c7a6a43b9d0b04618` |

---

## 5. Protocol for Future Delta Updates
When future revisions to code, templates, or business rules are committed to the upstream repository:
1. Ensure new changes are committed to Git with Conventional Commits (`feat(...)`, `fix(...)`, etc.).
2. Execute the companion master prompt: `MCP_TOOL_DELTA_EXTRACTOR.md`.
3. The delta extractor will:
   - Compare `v1.1.0` (or latest ledger tag) against `HEAD` via `git diff`.
   - Identify added/modified functions, template sheet layouts, and defect definitions.
   - Append an entry to this ledger.
   - Produce a delta migration patch for target MCP implementations.

