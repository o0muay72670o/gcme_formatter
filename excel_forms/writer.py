"""Deterministic workbook layout; AI never executes Excel commands."""
from copy import copy
from io import BytesIO
import math
from pathlib import Path
import re
import warnings
import unicodedata

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

LINE = Side(style="thin", color="000000")


def unique_name(wanted: str, existing: list[str]) -> str:
    base = re.sub(r"[\\/*?:\[\]]", "_", wanted).strip().strip("'")[:31] or "New"
    name, i = base, 2
    while name.casefold() in {s.casefold() for s in existing}:
        suffix = f"_{i}"
        name = base[:31 - len(suffix)] + suffix
        i += 1
    return name


def _rename_source(wb) -> str:
    old = wb.active
    if old is None:
        raise ValueError("Source workbook has no active worksheet")
    name = old.title
    # Renaming with openpyxl does not automatically repair dependent formulas.
    quoted, unquoted = f"'{name.replace(chr(39), chr(39)*2)}'!", f"{name}!"
    references = []
    for ws in wb:
        for cell in ws._cells.values():
            if cell.data_type == 'f' and isinstance(cell.value, str) and (quoted in cell.value or unquoted in cell.value):
                references.append(f"{ws.title}!{cell.coordinate}")
    defined = list(wb.defined_names.values())
    for ws in wb:
        defined.extend(ws.defined_names.values())
    if references or any(quoted in str(n.attr_text) or unquoted in str(n.attr_text) for n in defined):
        raise ValueError("The active source sheet has formula/named-range dependents. "
                         "Automatic rename is blocked to avoid breaking references; use Excel for this file.")
    new_name = unique_name('Old', [s for s in wb.sheetnames if s != name])
    old.title = new_name
    return new_name


def write_workbook(source: Path, destination: Path, plan: dict, extraction: dict, config: dict) -> dict:
    if destination.exists():
        raise FileExistsError(f"Output already exists: {destination}. Choose a different output folder.")
    if source.resolve() == destination.resolve():
        raise ValueError("Output must differ from input")
    issues = list(extraction.get("warnings", [])) + list(plan.get("review_notes", []))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        wb = load_workbook(source)
        master = load_workbook(config["master_file"])
    try:
        issues.extend(str(w.message) for w in caught)
        old_name = _rename_source(wb)
        name = unique_name(source.stem, wb.sheetnames)
        ws = wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        header = master[config["master_sheet"]]
        count, header_rows = config["last_column"], config["header_rows"]
        for row in header.iter_rows(min_row=1, max_row=header_rows, max_col=count):
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                if cell.data_type == 'f':
                    raise ValueError("Master header contains formulas; review dependencies before copying.")
                target = ws.cell(cell.row, cell.column, cell.value)
                target.font, target.fill, target.border = copy(cell.font), copy(cell.fill), copy(cell.border)
                target.alignment, target.protection = copy(cell.alignment), copy(cell.protection)
                target.number_format = cell.number_format
        for area in header.merged_cells.ranges:
            if area.max_row <= header_rows and area.max_col <= count:
                ws.merge_cells(str(area))
        for r in range(1, header_rows+1):
            ws.row_dimensions[r].height = header.row_dimensions[r].height or 15
        # OOXML width is based on the Normal font. Seven pixels is the baseline;
        # native Excel verification reports the actual measured physical width.
        width = config["column_width_cm"] * 96 / 2.54 / 7
        for col in range(1, count+1):
            letter = get_column_letter(col)
            ws.column_dimensions[letter] = ColumnDimension(ws, index=letter, width=width)
        index = {c["id"]: c for s in extraction["sheets"] for c in s["cells"]}
        # Pull typed header values from the source, not generated AI prose.
        for field in plan["header_fields"]:
            entries = [index[i] for i in field["source_ids"]]
            cell = ws[field["target"]]
            cell.value = entries[0]["value"] if len(entries) == 1 else " ".join(str(e["value"]) for e in entries)
            if len(entries) == 1:
                cell.number_format = entries[0]["format"]
            if cell.data_type == "f":
                cell.data_type = "s"
        title_cell = config.get('title_cell')
        if title_cell:
            cell = ws[title_cell]
            if isinstance(cell, MergedCell) or cell.row > header_rows or header[title_cell].value is not None:
                raise ValueError('Configured title_cell must be a blank top-left cell in the Master Header')
            if any(field['target'] == title_cell for field in plan['header_fields']):
                raise ValueError('title_cell conflicts with a header field mapping')
            cell.value = plan['title']
            if cell.data_type == 'f':
                cell.data_type = 's'

        font_size = config["body_font_size"]
        def box(row, start, end, text, bold=False, filled=False):
            for col in range(start, end+1):
                cell = ws.cell(row, col)
                cell.font = Font(name="Arial", size=font_size, bold=bold)
                cell.border = Border(top=LINE, bottom=LINE, left=LINE if col == start else Side(), right=LINE if col == end else Side())
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if filled:
                    cell.fill = PatternFill('solid', fgColor='E7E6E6')
            if end > start:
                ws.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)
            cell = ws.cell(row, start, text)
            if cell.data_type == 'f':
                cell.data_type = 's'  # source/form text is never executed as a formula
            pixels = (end-start+1)*config["column_width_cm"]*96/2.54
            per_line = max(1, (pixels-6)/(font_size*0.82))
            def display_length(text):
                return sum(3 if c == '\u2026' else 2 if unicodedata.east_asian_width(c) in ('W', 'F') else 1 for c in text)
            lines = sum(max(1, math.ceil(display_length(t)/per_line)) for t in str(text).split('\n'))
            ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 0, lines*(font_size+2)+3)

        r = header_rows+1
        for section in plan["sections"]:
            if section["title"]:
                box(r, 1, count, section["title"], bold=True, filled=True)
                r += 1
            columns = section["columns"]
            # Allocate more grid cells to long descriptions than response columns.
            samples = [[str(cell["text"]) for cell in row["cells"]] for row in section["rows"]]
            weights = [max(4, min(80, max([len(label)] + [len(row[i]) for row in samples]))) for i, label in enumerate(columns)]
            widths = [3 for _ in columns]
            for _ in range(count-sum(widths)):
                chosen = max(range(len(widths)), key=lambda i: weights[i]/widths[i])
                widths[chosen] += 1
            positions, start = [], 1
            for width in widths:
                positions.append((start, start+width-1))
                start += width
            if any(columns):
                for (start, end), text in zip(positions, columns):
                    box(r, start, end, text, bold=True, filled=True)
                r += 1
            for row in section["rows"]:
                for (start, end), cell in zip(positions, row["cells"]):
                    text = cell["text"]
                    if not text and cell["source_ids"]:
                        text = ' '.join(str(index[i]["value"]) for i in cell["source_ids"])
                    box(r, start, end, text)
                r += 1
        last_row = r-1
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = 'portrait'
        ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_margins = PageMargins(**{k:v/2.54 for k,v in config["margins_cm"].items()})
        ws.print_area = f'A1:{get_column_letter(count)}{last_row}'
        wb.active = wb.index(ws)
        height = sum(ws.row_dimensions[i].height or 15 for i in range(1, last_row+1))
        available = (29.7-config['margins_cm']['top']-config['margins_cm']['bottom'])*72/2.54
        scale = min(1.0, available/height)
        if font_size*scale < 7:
            issues.append(f"Estimated printed body font is {font_size*scale:.1f} pt. One-page layout needs review.")
        if any(s['images'] or s['charts'] for s in extraction['sheets']):
            issues.append('Source contains images/charts. Text extraction does not interpret them or copy them into the new body.')
        issues.append('openpyxl may not preserve unsupported Excel shapes or controls in the output copy; original input is unchanged.')
        destination.parent.mkdir(parents=True, exist_ok=True)
        # Exclusive creation protects outputs even if another run is writing.
        buffer = BytesIO()
        wb.save(buffer)
        with destination.open('xb') as stream:
            stream.write(buffer.getvalue())
        return {'output': str(destination.resolve()), 'new_sheet': name, 'renamed_sheet': old_name,
                'print_area': str(ws.print_area), 'estimated_scale': round(scale, 3),
                'native_print_verified': False, 'status': 'generated_needs_review', 'review_notes': issues}
    finally:
        wb.close()
        master.close()
