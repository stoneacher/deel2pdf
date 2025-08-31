#!/usr/bin/env python3
import os
import sys
import argparse
import pandas as pd
from fpdf import FPDF
from bs4 import BeautifulSoup

# Try to import Tk only if available (headless-safe).
try:
    from tkinter import Tk, filedialog, TclError  # type: ignore
except Exception:  # pragma: no cover
    Tk = None
    filedialog = None
    TclError = Exception

# -------------- Text Safety & Rendering Helpers --------------

def safe_cell(pdf, *args, **kwargs):
    """Wrapper for pdf.cell that reports failing content before raising."""
    try:
        pdf.cell(*args, **kwargs)
    except Exception as e:
        print("Failed on cell text:", args[2] if len(args) > 2 else args[0])
        raise e

def safe_multi_cell(pdf, *args, **kwargs):
    """Wrapper for pdf.multi_cell that reports failing content before raising."""
    try:
        pdf.multi_cell(*args, **kwargs)
    except Exception as e:
        print("Failed on multi_cell text:", args[0])
        raise e

def clean_text(text: object) -> str:
    """
    Keep only BMP (U+0000..U+FFFF) to avoid surrogate/emoji issues in classic FPDF.
    Replace non-BMP with the replacement character and log them for visibility.
    """
    if pd.isna(text):
        return "THERE IS NO TEXT!!!!"
    original = str(text)
    removed = [c for c in original if ord(c) > 0xFFFF]
    if removed:
        print(f"⚠️ Removed/marked non-BMP characters: {removed}")
    # Replace non-BMP with the replacement char so readers see something was there.
    filtered = ''.join(c if ord(c) <= 0xFFFF else "�" for c in original)
    return filtered

def render_list_item(pdf: FPDF, li, font_family: str, bullet="•", indent=0):
    """Render a single <li>, supporting nested <ul>/<ol>."""
    indent_str = "    " * indent
    content = []

    for child in li.contents:
        if isinstance(child, str):
            s = child.strip()
            if s:
                content.append(clean_text(s))
        elif getattr(child, "name", None) in ["p", "span", None]:
            txt = child.get_text(strip=True)
            if txt:
                content.append(clean_text(txt))
        elif getattr(child, "name", None) in ["ul", "ol"]:
            sub_bullet = "•" if child.name == "ul" else "1."
            for j, sub_li in enumerate(child.find_all("li", recursive=False), 1):
                render_list_item(
                    pdf,
                    sub_li,
                    font_family=font_family,
                    bullet=sub_bullet if child.name == "ul" else f"{j}.",
                    indent=indent + 1
                )
            continue
        else:
            txt = child.get_text(strip=True)
            if txt:
                content.append(clean_text(txt))

    full_line = f"{indent_str}{bullet} {' '.join(content).strip()}".rstrip()
    if full_line:
        pdf.set_font(font_family, size=10)
        safe_multi_cell(pdf, 0, 6, full_line, align="L")

def render_html_comment(pdf: FPDF, html_content, font_family: str):
    """
    Render a small subset of HTML into the PDF:
    - p, strong/b, em/i, u, ul/ol/li
    - Tables are rejected (unsupported).
    """
    if pd.isna(html_content) or not str(html_content).strip():
        pdf.set_font(font_family, style="I", size=10)
        safe_multi_cell(pdf, 0, 6, "No comment provided.", align="L")
        return

    soup = BeautifulSoup(html_content, "html.parser")

    if soup.find("table"):
        print("Error: HTML contains a table. PDF generation for tables is not supported.")
        return

    def render_tag(tag):
        # Text node
        if getattr(tag, "name", None) is None:
            txt = clean_text(tag)
            if txt.strip():
                pdf.set_font(font_family, size=10)
                safe_multi_cell(pdf, 0, 6, txt, align="L")
            return

        # Paragraph
        if tag.name == "p":
            for inner in tag.children:
                render_tag(inner)
            pdf.ln(2)
            return

        # Strong/Bold
        if tag.name in ["strong", "b"]:
            txt = clean_text(tag.get_text())
            if txt.strip():
                pdf.set_font(font_family, style="B", size=10)
                safe_multi_cell(pdf, 0, 6, txt, align="L")
            return

        # Emphasis/Italic
        if tag.name in ["em", "i"]:
            txt = clean_text(tag.get_text())
            if txt.strip():
                pdf.set_font(font_family, style="I", size=10)
                safe_multi_cell(pdf, 0, 6, txt, align="L")
            return

        # Underline (real underline)
        if tag.name == "u":
            txt = clean_text(tag.get_text())
            if txt.strip():
                pdf.set_font(font_family, style="U", size=10)
                safe_multi_cell(pdf, 0, 6, txt, align="L")
            return

        # Unordered list
        if tag.name == "ul":
            for li in tag.find_all("li", recursive=False):
                render_list_item(pdf, li, font_family=font_family, bullet="•", indent=0)
            return

        # Ordered list
        if tag.name == "ol":
            for i, li in enumerate(tag.find_all("li", recursive=False), 1):
                render_list_item(pdf, li, font_family=font_family, bullet=f"{i}.", indent=0)
            return

        # Fallback: dump text
        txt = clean_text(tag.get_text())
        if txt.strip():
            pdf.set_font(font_family, size=10)
            safe_multi_cell(pdf, 0, 6, txt, align="L")

    for child in soup.children:
        render_tag(child)

# -------------- Font Handling --------------

FONT_PRESETS = {
    # Default preset: Google Noto Sans
    "noto": {
        "family_name": "NotoSans",
        "folder": "NotoSans",
        "files": {
            "regular": "NotoSans-Regular.ttf",
            "bold": "NotoSans-Bold.ttf",
            "italic": "NotoSans-Italic.ttf",
            "bold_italic": "NotoSans-BoldItalic.ttf",
        },
    },
    # Legacy DejaVu Sans preset (was used for first versions)
    "dejavu": {
        "family_name": "DejaVuSans",
        "folder": "DejaVuSans",
        "files": {
            "regular": "DejaVuSans.ttf",
            "bold": "DejaVuSans-Bold.ttf",
            "italic": "DejaVuSans-Oblique.ttf",
            "bold_italic": "DejaVuSans-BoldOblique.ttf",
        },
    },
}

def ensure_fonts(base_fonts_dir: str, preset_key: str) -> tuple[str, dict]:
    """
    Verify font files exist for the selected preset and return:
    - family_name used in FPDF (e.g., "NotoSans")
    - dict of absolute file paths for styles
    """
    if preset_key not in FONT_PRESETS:
        print(f"Error: Unknown font preset '{preset_key}'. Available: {list(FONT_PRESETS.keys())}")
        sys.exit(1)

    preset = FONT_PRESETS[preset_key]
    family_name = preset["family_name"]
    folder_name = preset["folder"]

    font_dir = os.path.join(base_fonts_dir, folder_name)
    files = preset["files"]

    paths = {}
    for style, filename in files.items():
        font_path = os.path.join(font_dir, filename)
        if not os.path.exists(font_path):
            print(f"Error: Font file '{filename}' not found.")
            print(f"Expected at: {font_path}")
            print("Tip: Place the TTFs under 'fonts/{folder}/' as documented.".format(folder=folder_name))
            sys.exit(1)
        paths[style] = font_path

    return family_name, paths

def add_fonts(pdf: FPDF, family_name: str, fonts: dict):
    """Register fonts with FPDF for Unicode output."""
    pdf.add_font(family_name, "", fonts["regular"], uni=True)
    pdf.add_font(family_name, "B", fonts["bold"], uni=True)
    pdf.add_font(family_name, "I", fonts["italic"], uni=True)
    pdf.add_font(family_name, "BI", fonts["bold_italic"], uni=True)

# -------------- Core Generation Logic --------------

def detect_reviewer_column(df: pd.DataFrame) -> str:
    """Detect the reviewer name column with a couple of common variants."""
    if "Reviewer's name" in df.columns:
        return "Reviewer's name"
    if "Reviewers name" in df.columns:
        return "Reviewers name"
    raise KeyError("Reviewer name column not found (expected one of: \"Reviewer's name\", \"Reviewers name\").")

def choose_file_interactively() -> str:
    """Open a file dialog to select the Excel file."""
    if Tk is None or filedialog is None:
        return ""  # Not available in this environment
    try:
        root = Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title="Select the Excel file",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        root.update_idletasks()
        root.destroy()
        return file_path or ""
    except TclError:
        # Headless or display error
        return ""

def generate_pdfs(xlsx_path: str, font_preset: str = "noto"):
    if not xlsx_path:
        print("No file selected. Exiting...")
        sys.exit(0)

    print(f"Selected file: {xlsx_path}")

    output_dir = os.path.join(os.path.dirname(xlsx_path), "generated_pdfs")
    os.makedirs(output_dir, exist_ok=True)

    xls = pd.ExcelFile(xlsx_path)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

    # Map internal feedback types to display headers
    feedback_type_mapping = {
        "self_shared_feedback": "Employee Review",
        "auto_shared_feedback": "Supervisor Review",
        "shared_feedback": "Supervisor Review"
    }

    # Grouping columns
    try:
        reviewer_col = detect_reviewer_column(df)
    except KeyError as e:
        print(f"Error: {e}")
        sys.exit(1)

    group_columns = ["Reviewee name", "Review Cycle name", "Team - Reviewee", "Position - Reviewee", reviewer_col]
    missing = [c for c in group_columns if c not in df.columns]
    if missing:
        print(f"Error: Missing expected columns: {missing}")
        sys.exit(1)

    grouped = df.groupby(group_columns)

    # Fonts
    script_dir = os.path.dirname(os.path.abspath(__file__))
    base_fonts_dir = os.path.join(script_dir, "fonts")
    font_family, fonts = ensure_fonts(base_fonts_dir, font_preset)

    # Generate one PDF per group
    for (reviewee, review_cycle, reviewee_team, reviewee_position, reviewer), group in grouped:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        add_fonts(pdf, font_family, fonts)

        # Title block
        pdf.set_font(font_family, style="B", size=16)
        safe_cell(pdf, 200, 10, clean_text(f"{reviewee} - {review_cycle}"), ln=True, align="C")
        pdf.set_font(font_family, size=10)
        safe_cell(pdf, 0, 5, clean_text(f"Team: {reviewee_team}"), ln=True, align="C")
        safe_cell(pdf, 0, 5, clean_text(f"Position: {reviewee_position}"), ln=True, align="C")
        pdf.set_font(font_family, style="BI", size=10)
        pdf.ln(10)

        # Sections
        for feedback_type, header in feedback_type_mapping.items():
            feedback_data = group[group["Feedback type"] == feedback_type]

            if not feedback_data.empty:
                pdf.set_font(font_family, size=14)
                safe_cell(pdf, 200, 10, clean_text(header), ln=True)
                pdf.ln(5)

                for _, row in feedback_data.iterrows():
                    # Date / reviewer
                    pdf.set_font(font_family, size=8)
                    try:
                        date_str = pd.to_datetime(row['Review cycle launch date']).strftime('%Y-%m-%d')
                    except Exception:
                        date_str = clean_text(str(row.get('Review cycle launch date', '')))
                    safe_multi_cell(pdf, 0, 8, clean_text(f"Date: {date_str} / Reviewer: {reviewer}"), align="L")

                    # Question
                    pdf.set_font(font_family, style="B", size=12)
                    safe_multi_cell(pdf, 0, 7, clean_text(f"Question: {row['Question']}"), align="L")

                    # Description
                    pdf.set_font(font_family, style="B", size=10)
                    safe_multi_cell(pdf, 0, 6, clean_text(f"{row['Question description']}"), align="L")

                    # Comment (HTML)
                    pdf.set_font(font_family, style="B", size=10)
                    safe_cell(pdf, 0, 6, "Response:", ln=True)
                    render_html_comment(pdf, row['Response comment'], font_family=font_family)
                    pdf.ln(5)

        sanitized_name = str(reviewee).replace(" ", "_")
        filename = f"{sanitized_name}_{review_cycle}.pdf"
        filepath = os.path.join(output_dir, filename)

        try:
            pdf.output(filepath)
        except Exception as e:
            print("\nError during pdf.output")
            print(f"Failed to write: {filepath}")
            raise e

        print(f"Generated: {filepath}")

    print(f"\nAll PDFs have been created in: '{output_dir}'")

# -------------- CLI & Entry Point --------------

def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Generate review PDFs from an Excel export."
    )
    p.add_argument(
        "-f", "--file",
        dest="file",
        help="Path to the Excel file (.xlsx or .xls). If omitted, a file dialog is shown (if available).",
    )
    p.add_argument(
        "--font",
        dest="font",
        choices=sorted(FONT_PRESETS.keys()),
        default="noto",
        help="Font preset to use: 'noto' (default) or 'dejavu'.",
    )
    return p.parse_args(argv)

def main(argv=None):
    args = parse_args(argv)
    xlsx_path = args.file

    # If no path provided via CLI, try a file dialog (if available)
    if not xlsx_path:
        xlsx_path = choose_file_interactively()

    # Basic validation
    if xlsx_path and not os.path.isfile(xlsx_path):
        print(f"Error: File not found: {xlsx_path}")
        sys.exit(1)

    generate_pdfs(xlsx_path, font_preset=args.font)

if __name__ == "__main__":
    main()
