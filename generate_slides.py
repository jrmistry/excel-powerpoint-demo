#!/usr/bin/env python3
"""
Excel → PowerPoint slide generator.

For each tab in the Excel file, one slide is produced from the template.
- The text  {{name}}  anywhere in the slide is replaced with the tab name.
- Rows from each tab are appended to the table whose column headers match
  the Excel column headers exactly (non-matching columns are skipped).

Usage:
    python generate_slides.py [excel_file] [template_file] [output_file]

Defaults:
    excel_file    = data.xlsx
    template_file = template.pptx
    output_file   = output.pptx
"""

import copy
import sys
from pathlib import Path

import openpyxl
from lxml import etree
from pptx import Presentation


# ── configuration ─────────────────────────────────────────────────────────────
EXCEL_FILE    = "data.xlsx"
TEMPLATE_FILE = "template.pptx"
OUTPUT_FILE   = "output.pptx"

# Text in the slide that gets replaced with the Excel tab name.
NAME_PLACEHOLDER = "{{name}}"
# ──────────────────────────────────────────────────────────────────────────────


NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


# ── helpers ───────────────────────────────────────────────────────────────────

def get_table(slide):
    """Return the first table shape found on *slide*, or None."""
    for shape in slide.shapes:
        if shape.has_table:
            return shape.table
    return None


def cell_text(cell):
    """Return stripped text content of a table cell."""
    return cell.text_frame.text.strip()


def append_data_row(table, col_map, row_data):
    """
    Build a brand-new <a:tr> element and append it to *table*.

    Rows are constructed from scratch rather than copied from an existing row.
    Copying any existing row (including the header) inherits cell properties —
    fill colours, run-formatting flags, and sometimes implicit lock attributes —
    that cause PowerPoint to render the pasted cells as non-editable.
    A clean minimal row avoids all of that; the table's built-in style still
    applies banded-row colours automatically based on each row's position.

    col_map  : {pptx_column_index: excel_column_name}
    row_data : {excel_column_name: value}
    """
    tbl = table._tbl
    header_tr = tbl.findall(f"{{{NS}}}tr")[0]

    # Borrow row height from the header so sizing stays consistent.
    row_height = header_tr.get("w", "370840")
    num_cols   = len(header_tr.findall(f"{{{NS}}}tc"))

    new_tr = etree.Element(f"{{{NS}}}tr")
    new_tr.set("w", row_height)

    for col_idx in range(num_cols):
        col_name = col_map.get(col_idx)
        if col_name is not None:
            val  = row_data.get(col_name)
            text = "" if val is None else str(val)
        else:
            text = ""

        tc     = etree.SubElement(new_tr, f"{{{NS}}}tc")
        txBody = etree.SubElement(tc, f"{{{NS}}}txBody")
        etree.SubElement(txBody, f"{{{NS}}}bodyPr")
        etree.SubElement(txBody, f"{{{NS}}}lstStyle")
        p      = etree.SubElement(txBody, f"{{{NS}}}p")

        if text:
            r   = etree.SubElement(p, f"{{{NS}}}r")
            rPr = etree.SubElement(r, f"{{{NS}}}rPr")
            rPr.set("lang", "en-US")
            rPr.set("dirty", "0")
            t   = etree.SubElement(r, f"{{{NS}}}t")
            t.text = text

        etree.SubElement(tc, f"{{{NS}}}tcPr")

    tbl.append(new_tr)


def replace_placeholder(slide, placeholder, value):
    """Replace every occurrence of *placeholder* in all text runs on *slide*."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)


def make_slide_from_template(prs, original_sp_tree, layout):
    """
    Add a new slide to *prs* and populate it with a fresh copy of
    *original_sp_tree* (the shape tree snapshotted from the template slide).
    """
    new_slide = prs.slides.add_slide(layout)
    sp_tree = new_slide.shapes._spTree

    # Remove the auto-generated placeholder shapes (keep first 2 fixed children:
    # nvGrpSpPr and grpSpPr).
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)

    # Paste a fresh copy of the original template shapes.
    for child in list(original_sp_tree)[2:]:
        sp_tree.append(copy.deepcopy(child))

    return new_slide


# ── main logic ────────────────────────────────────────────────────────────────

def process(excel_path, template_path, output_path, placeholder=NAME_PLACEHOLDER):
    wb = openpyxl.load_workbook(excel_path)
    prs = Presentation(template_path)

    if not prs.slides:
        sys.exit("Error: template contains no slides.")

    # Snapshot the original template slide *before* we mutate anything.
    tmpl_slide = prs.slides[0]
    original_sp_tree = copy.deepcopy(tmpl_slide.shapes._spTree)
    tmpl_layout = tmpl_slide.slide_layout

    sheets = wb.sheetnames
    print(f"Sheets found: {', '.join(sheets)}\n")

    slides_created = []

    for idx, sheet_name in enumerate(sheets):
        ws = wb[sheet_name]

        # Read column headers from row 1 (stop at first blank cell).
        headers = []
        for cell in ws[1]:
            if cell.value is None:
                break
            headers.append(str(cell.value).strip())

        if not headers:
            print(f"[SKIP] '{sheet_name}': no headers found in row 1.")
            continue

        # Read data rows (skip fully empty rows).
        data_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in row):
                continue
            data_rows.append({
                headers[c]: row[c]
                for c in range(min(len(headers), len(row)))
            })

        print(f"Sheet '{sheet_name}': {len(data_rows)} data row(s), "
              f"columns: {headers}")

        # Get or create the slide for this sheet.
        if idx == 0:
            slide = prs.slides[0]   # reuse the template slide directly
        else:
            slide = make_slide_from_template(prs, original_sp_tree, tmpl_layout)

        # Inject the tab name.
        replace_placeholder(slide, placeholder, sheet_name)

        # Locate the table.
        table = get_table(slide)
        if table is None:
            print(f"  [WARN] No table found on slide — skipping '{sheet_name}'.")
            continue

        # Map: PowerPoint column index → Excel column name (exact-match only).
        pptx_cols = [
            cell_text(table.rows[0].cells[c])
            for c in range(len(table.rows[0].cells))
        ]
        col_map = {c: name for c, name in enumerate(pptx_cols) if name in headers}

        if not col_map:
            print(f"  [WARN] No column names match between Excel and the table.")
            print(f"         Table headers : {pptx_cols}")
            print(f"         Excel headers : {headers}")
            continue

        matched = [col_map[c] for c in sorted(col_map)]
        skipped = [h for h in headers if h not in matched]
        print(f"  Matched : {matched}")
        if skipped:
            print(f"  Skipped : {skipped}")

        # Remove any pre-existing non-header rows (template placeholders) so
        # only the header row remains before we insert real data.
        tbl = table._tbl
        existing_trs = tbl.findall(f"{{{NS}}}tr")
        for tr in existing_trs[1:]:
            tbl.remove(tr)

        # Append one table row per Excel data row.
        for row_data in data_rows:
            append_data_row(table, col_map, row_data)

        slides_created.append(sheet_name)

    prs.save(output_path)
    print(f"\nDone — {len(slides_created)} slide(s) created: {', '.join(slides_created)}")
    print(f"Output saved to: {output_path}")


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    args = sys.argv[1:]
    excel    = args[0] if len(args) > 0 else EXCEL_FILE
    template = args[1] if len(args) > 1 else TEMPLATE_FILE
    output   = args[2] if len(args) > 2 else OUTPUT_FILE

    for path, label in [(excel, "Excel"), (template, "Template")]:
        if not Path(path).exists():
            sys.exit(f"Error: {label} file not found: '{path}'")

    process(excel, template, output)
