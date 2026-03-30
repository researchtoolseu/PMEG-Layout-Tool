# PMEG Layout Tool

This folder contains a clinically oriented, **true-scale (1:1)** digital planning tool for **Physician-Modified Endografts (PMEG)**.

The tool generates printable templates that can be cut, rolled, and used directly on the back table to mark fenestrations accurately on the graft fabric.

Everything is referenced to the **nominal graft diameter (device)** - **not** to the native aortic diameter.

---

## Files

* **pmeg_layout_mm_paper_v2.10_xlsx_last_output_folder.py**  
  Main Python script

* **PMEG_Input_Template_with_Run_Button.xlsx**  
  Recommended Excel input template with structured metadata and run button

* **targets_v2.10.csv** *(optional / legacy)*  
  CSV input format

* **README_v2.10.md**  
  Documentation

---

## What the script does

* Unrolls the graft circumference using the nominal graft diameter
* Places fenestrations using angular (`theta_deg`) and longitudinal data
* Produces **true-scale output (1 mm = 1 mm)** on A4/A3 paper
* Carries the **physician_name** metadata through the workflow for traceability

---

## Outputs

### Main PDF / PNG

Includes:

* millimeter grid
* fenestration markers and labels
* wrap-edge lines (true graft perimeter)
* cut guides
* clock-face orientation
* AP orientation markers (12 & 6 o'clock)
* check marker (`✓`) to prevent 180-degree rotation
* reduction-tie planning guides
* 100 x 100 mm calibration square
* dual longitudinal measurement scales:
  * center-to-center
  * bottom-to-bottom
* patient metadata
* **measurements by / physician_name** line when provided

### Film PDF (transparent film / punch card)

Minimal geometry for back-table use:

* graft boundaries
* fenestrations
* AP marker
* reduction-tie guides
* check marker (`✓`)
* calibration square
* optional small physician traceability line when `physician_name` is present

### Report files

Generated automatically:

* `*_REPORT.txt` - human-readable summary
* `*_REPORT.csv` - structured data for reuse

Includes:

* all metadata
* **physician_name**
* fenestration coordinates
* center-to-center distances
* bottom-to-bottom distances
* anchor-to-target distances

In packaged / executable use, the tool also writes a small sidecar file:

* `last_output_folder.txt`

This file is created alongside the executable and contains the **full path of the most recently created output folder**. It allows the Excel/VBA launcher to open the exact newly generated patient run folder after the app finishes.

---

## New traceability metadata

A new metadata field is supported:

```text
physician_name
```

Use it to record the physician who performed the measurements / planning for that case.

Why this matters:

* supports multi-physician departmental workflows
* improves traceability and auditability
* helps with inter-observer comparison and research datasets
* ensures the responsible physician is visible in the generated outputs

For backward compatibility, the script also accepts these aliases and normalizes them to `physician_name`:

* `planner_name`
* `measuring_physician`
* `measured_by`

---

## Requirements

### Python

* Python **3.10 - 3.12**

### Install dependencies

```bash
python -m pip install matplotlib openpyxl
```

(macOS: use `python3`)

---

## Input methods

### Recommended: Excel (.xlsx)

The Excel file contains structured fields for metadata and targets.

Metadata fields include:

* patient_name
* patient_age
* study_date
* **physician_name**
* graft_diam_mm
* paper
* orientation
* film_height_mm
* tie settings
* cut margin (`cut_margin_mm`)
* anchor settings

Target fields include:

* Name
* Theta_deg
* Fen_diam_mm
* Dist_from_zero_mm
* Y_mm
* Notes

### CSV (legacy support)

Still supported for flexibility.

Metadata lines at the top of the CSV can include:

```text
# patient_name: John Doe
# patient_age: 68
# study_date: 2026-02-01
# physician_name: Dr Jane Smith
# graft_diam_mm: 30
```

---

## Core concept

### Planning mode (recommended workflow)

* **Anchor** = most proximal fenestration (usually CA)
* **ZERO reference** = **bottom of anchor fenestration**
* Input:

```text
dist_from_zero_mm = bottom(anchor) -> bottom(target)
```

Example:

```text
CA -> SMA = 18 mm (bottom-to-bottom)
```

The script automatically converts bottom-to-bottom distances to centerline coordinates for plotting.

---

## Measurement scales on output

### Scale A - Center-to-center

Between adjacent fenestration centers

### Scale B - Bottom-to-bottom

Between adjacent fenestrations, plus:

* TOP -> bottom of first fenestration

---

## Automatic patient folder system

If `--out` is **not** specified, outputs are stored as:

```text
/Patients/
    YYYY-MM-DD_PatientName/
        v001_timestamp/
        v002_timestamp/
```

Each run creates a new version folder and stores:

* PDFs
* PNG
* report files
* original input file (traceability)

Additionally, the app writes `last_output_folder.txt` next to the executable / script, containing the path of that exact run folder. This is useful when Excel should open the **specific newly created patient folder** instead of the generic `Patients` root.

---

## How to run

### Windows

```bash
cd C:\Path\To\PMEG
python pmeg_layout_mm_paper_v2.10_xlsx_last_output_folder.py --input PMEG_Input_Template_with_Run_Button.xlsx
```

### macOS

```bash
cd /Users/yourname/Path/To/PMEG
python3 pmeg_layout_mm_paper_v2.10_xlsx_last_output_folder.py --input PMEG_Input_Template_with_Run_Button.xlsx
```

---

## Printing (critical)

* Print at **100% / Actual size**
* Disable scaling
* Verify the calibration square

If incorrect -> **DO NOT USE**

---

## Safety features

* AP marker (12-6 o'clock)
* Check marker (`✓`) to prevent rotation errors
* Wrap edges clearly marked
* Warnings for out-of-bounds fenestrations or incorrect geometry

---

## Quick checklist

* Anchor defined correctly
* Distances are bottom-to-bottom
* `theta_deg` values are correct
* `physician_name` entered when measurements are assigned to a specific physician
* Calibration square verified
* No scaling in printing

---

## Disclaimer

This tool supports planning and documentation only.

Clinical judgment and responsibility remain entirely with the treating physician.

---

## Credits

Created by:

* **Michael A. Lazaris**
* **Andreas M. Lazaris**

Tools:

* Python
* Matplotlib
* OpenPyXL
* ChatGPT (OpenAI)

© 2026
