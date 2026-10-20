# OMR Sheet Scanner

A desktop application that automatically reads scanned OMR (Optical Mark Recognition) answer sheets, grades them against an Excel answer key, and exports everything to a formatted Excel report. Built with Python, OpenCV and Tkinter during my internship.

## Features

- **Batch scanning**: select a folder and process all images at once (`.jpg`, `.png`, `.bmp`, `.tif`)
- **Auto-correction**: detects registration squares, fixes rotation/perspective, and flips upside-down sheets
- **Bubble detection**: reads Sr. No., Grade, District, Area, Paper, CRP No., School Code, Gender and 60 MCQ answers
- **Smart validation**: blank or multiple-marked bubbles are flagged and receive 0 marks
- **Auto grading**: compares answers with an Excel answer key (Q1-Q60, A-D)
- **Excel report**: per-question T/F with green/red highlighting, total marks, obtained marks, percentage and status
- **Live GUI**: progress bar, processed-sheet preview and processing log; scanning runs in a background thread so the UI never freezes

## How It Works

1. Image is loaded and rotated to portrait
2. Four corner registration squares are detected with contour analysis
3. Perspective transform maps the sheet to a fixed 2480 x 3507 template
4. Image is binarised (Otsu threshold) and the fill ratio of every bubble is measured
5. Answers are compared with the answer key and the results are written to Excel

## Tech Stack

Python | OpenCV | NumPy | openpyxl | Pillow | Tkinter

## Download

Windows app (no Python needed): [Download latest release](https://github.com/AqsaImran-IT/omr-sheet-scanner/releases/latest)

## Installation

```bash
git clone https://github.com/AqsaImran-IT/omr-sheet-scanner.git
cd omr-sheet-scanner
pip install -r requirements.txt
```

## Usage

```bash
python main.py
```

1. Click **Browse Source Folder** and select the folder with scanned OMR images
2. Click **Browse Answer Key** and select your Excel file
3. Click **Start Scanning**
4. Results are saved to `output/OMR_Scan_Results.xlsx`, and the corrected sheets are saved in `output/processed_images/`

### Answer key format

| Column A | Column B |
|----------|----------|
| 1        | A        |
| 2        | C        |
| ...      | ...      |
| 60       | D        |

## Project Structure

```
omr-sheet-scanner/
├── main.py            # Scanner logic + GUI
├── requirements.txt
├── .gitignore
└── README.md
```

## Notes

- Bubble coordinates are calibrated for one specific answer-sheet template (2480 x 3507 px). A different sheet layout needs updated coordinates.
- Handwritten fields (student name, school name) are left blank for manual verification.

## License

MIT License. See the LICENSE file for details.

## Author

**Aqsa Imran**: [LinkedIn](https://linkedin.com/in/aqsa-imran-57336534b) | [GitHub](https://github.com/AqsaImran-IT)
