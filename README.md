# Deel Review Export → PDF Generator

This tool converts **[Deel](https://www.deel.com/)** review exports (Excel format) into clean, readable **PDF reports** for each employee.
It’s designed to make Deel’s raw exports more shareable and presentation-friendly.

By default, the script uses [**Noto Sans**](https://fonts.google.com/noto/specimen/Noto+Sans) for broad Unicode coverage, with optional support for **DejaVu Sans**.

![example](assets/Screenshot%202025-08-31%20at%2022.05.30.png)


Here’s the polished **README.md** with the **tkinter dependency note** integrated smoothly:

---

## Features

* Works directly with **Excel exports from Deel** review cycles.
* Generates **one PDF per employee**, grouped by:

  * Reviewee name
  * Review cycle
  * Team
  * Position
  * Reviewer
* Preserves formatting for HTML in comments:

  * Paragraphs
  * Bold / italic / underline
  * Lists (`<ul>`, `<ol>`)
* Safe Unicode handling (unsupported emoji replaced with `�`).
* Configurable font family (Noto Sans or DejaVu Sans).
* Works in GUI (file picker) or CLI/headless environments.

---

## Installation

Clone the repo and install requirements:

```bash
git clone https://github.com/your-username/deel2pdf.git
cd deel2pdf
pip install -r requirements.txt
```

---

## GUI File Picker (Optional Dependency)

If you run the script *without* `--file`, it will try to open a file picker using **tkinter**.

* **Windows/macOS** → included in most Python installers.
* **Linux** → may need to install separately:

  ```bash
  sudo apt-get install python3-tk
  ```
* **Headless/servers** → just pass the file explicitly with `--file /path/to/export.xlsx` and `tkinter` is not needed.

---

## Fonts

Font files are required to embed text into the PDFs.

### Noto Sans (default, recommended)

Download from Google Fonts:
👉 [Noto Sans download](https://fonts.google.com/download?family=Noto%20Sans)

Put them in:

```
fonts/NotoSans/
  ├─ NotoSans-Regular.ttf
  ├─ NotoSans-Bold.ttf
  ├─ NotoSans-Italic.ttf
  └─ NotoSans-BoldItalic.ttf
```

### DejaVu Sans (legacy, optional)

Download from [DejaVu project](https://dejavu-fonts.github.io/) and place in:

```
fonts/DejaVuSans/
  ├─ DejaVuSans.ttf
  ├─ DejaVuSans-Bold.ttf
  ├─ DejaVuSans-Oblique.ttf
  └─ DejaVuSans-BoldOblique.ttf
```

---

## Usage

### 1. Export from Deel

* In Deel, go to **Reviews → Export results (Excel)**.

### 2. Run the script

#### With file picker (desktop only)

```bash
python deel2pdf.py
```

#### With explicit file path

```bash
python deel2pdf.py --file example_export.xlsx
```

#### Choose font preset

```bash
# Use Noto Sans (default)
python deel2pdf.py --file example_export.xlsx --font noto

# Use DejaVu Sans
python deel2pdf.py --file example_export.xlsx --font dejavu
```

---

## ⚠️ Notes & Limitations

* Emoji and other non-BMP characters are replaced with `�`.
* `<table>` tags in comments are not supported.
* For **CJK** or **RTL scripts**, use the appropriate [Noto family](https://fonts.google.com/noto) instead of the default Latin set.

---

## License

* Script: MIT License
* Fonts:

  * **Noto Sans** → [SIL Open Font License](https://openfontlicense.org/)
  * **DejaVu Sans** → [Bitstream Vera License](https://dejavu-fonts.github.io/)

