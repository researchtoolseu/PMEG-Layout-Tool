# PMEG Layout Tool

## Overview

The **PMEG Layout Tool** is a clinically oriented, **true-scale (1:1)** digital planning tool for **Physician-Modified Endografts (PMEG)**. 

It generates printable templates that can be cut, rolled, and used directly on the back table to accurately mark fenestrations on the graft fabric.

All calculations are based on the **nominal graft diameter (device)** — **not** the native aortic diameter.

The tool is designed to support precision, reproducibility, and traceability in PMEG planning workflows.

---

## Main Features

* True-scale output (**1 mm = 1 mm**) on A4/A3 paper
* Graft unrolling based on nominal diameter (π·D)
* Fenestration positioning using:

  * angular position (`theta_deg`)
  * longitudinal distances
* Dual longitudinal measurement scales:

  * center-to-center distances
  * bottom-to-bottom distances
* Anchor-based planning mode (recommended workflow)
* Automatic conversion from bottom-based measurements to center coordinates
* Reduction tie planning guides (**configurable per case**)
* AP orientation markers (12–6 o’clock)
* Anti-rotation check marker (✓)
* Wrap edges clearly defined (true graft circumference)
* Cut guides with configurable margins
* Physician traceability via `physician_name` metadata
* Automatic patient-specific folder creation with versioning
* Report generation:

  * TXT (human-readable)
  * CSV (structured data for research/reuse)

---

## Input Format

### Recommended: Excel (.xlsm)

The Excel input file provides a structured and user-friendly interface.

### Metadata fields

* patient_name
* patient_age
* study_date
* physician_name
* graft_diam_mm
* paper (A4 / A3)
* orientation (portrait / landscape)
* film_height_mm
* tie_num_rows
* tie_edge_pad_mm
* tie_positions_clock *(NEW in v2.12)*
* cut_margin_mm
* ap_anchor
* v_anchor
* proximal_cover_mm
* anchor

### Target fields

* name
* theta_deg *(signed angle: right = positive, left = negative)*
* fen_diam_mm
* dist_from_zero_mm *(planning mode)*
* y_mm *(optional / legacy mode)*
* notes

![INPUT_FORMAT](images/INPUT_FORMAT.png)

---

### CSV (legacy support)

Metadata lines are defined using `#`:

```
# patient_name: John Doe
# patient_age: 68
# study_date: 2026-02-01
# physician_name: Dr Jane Smith
# graft_diam_mm: 30
# tie_positions_clock: 5,6,7
```

---

## Output Files

Each run generates a **patient-specific folder**:

```
Patients/
  YYYY-MM-DD_PatientName/
    v001_timestamp/
    v002_timestamp/
```

### Main outputs

* PDF (true-scale layout)
* PNG image

Includes:

* millimeter grid
* fenestration markers
* wrap edges
* cut guides
* clock-face orientation
* AP markers
* check marker (✓)
* reduction tie guides
* calibration square (100 × 100 mm)
* measurement scales
* patient metadata

![PMEG_LAYOUT](images/PMEG_LAYOUT.png)

---

### Film output (transparent film)

* simplified geometry for back-table use
* includes:

  * graft boundaries
  * fenestrations
  * AP marker
  * tie guides
  * calibration square

![PMEG_LAYOUT_FILM](images/PMEG_LAYOUT_FILM.png)

---

### Reports

* `*_REPORT.txt`
* `*_REPORT.csv`

Include:

* all metadata
* physician_name
* fenestration coordinates
* distances:

  * center-to-center
  * bottom-to-bottom
  * anchor-to-target

---

### Additional files

* copy of input file (traceability)
* `last_output_folder.txt` (used by Excel to open latest output)

---

## How to Run

### Python

```bash
python pmeg_layout_tool_v2.12.py --input PMEG_Input.xlsm
```

(macOS: use `python3`)

---

### Executable (recommended for users)

1. Open Excel input file
2. Enable macros
3. Click **Run PMEG**
4. Confirm execution
5. Wait until completion
6. Output folder opens automatically

---

## Versioning

The tool follows **version-based naming**:

* Script: `pmeg_layout_tool_v2.12.py`
* Current version: **v2.12**

Each version introduces incremental improvements and is documented separately.

Outputs are also versioned per patient (v001, v002, etc.), ensuring full traceability.

---

## Notes / Printing Instructions (critical)

* Print at **100% / Actual Size**
* Disable scaling / fit-to-page
* Verify the **100 × 100 mm calibration square**

If calibration is incorrect → **DO NOT USE**

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

## Changelog

A detailed version history is available in:

```
changelog.txt
```

---

## Disclaimer

This tool supports planning and documentation only.

Users and treating physicians are fully responsible for:

* measurement accuracy
* correct data entry
* clinical interpretation
* procedural execution

The tool does not replace clinical judgment.

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

---

## Usage Notice

The PMEG Layout Tool is provided for **academic, educational, and research purposes only**.

Commercial use, redistribution, or integration into commercial products is **not permitted** without prior written permission from the authors.

For licensing inquiries, please contact:
andreaslazaris@hotmail.com

---

## License

This project is licensed under the **Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0)**.

Commercial use is not permitted without prior permission.

See the LICENSE file for details.

© 2026




