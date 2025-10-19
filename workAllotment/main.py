import csv
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def read_work_csv(csv_path):
    """Read the work CSV file and return all rows"""
    with open(csv_path, "r", encoding="utf-8") as file:
        reader = csv.reader(file)
        rows = list(reader)
    return rows


def parse_names(name_string):
    """Parse comma-separated names and return a list of cleaned names"""
    if not name_string or name_string.strip() == "":
        return []

    # Split by comma and clean up whitespace
    names = [name.strip() for name in name_string.split(",")]
    # Filter out empty strings
    names = [name for name in names if name]
    return names


def is_empty_row(row):
    """Check if a row is essentially empty (all fields empty or whitespace)"""
    return all(cell.strip() == "" for cell in row)


def get_non_empty_columns(header, rows):
    """Identify which columns have at least one non-empty value"""
    num_cols = len(header)
    has_content = [False] * num_cols

    for row in rows:
        if is_empty_row(row):
            continue
        for i in range(min(len(row), num_cols)):
            if row[i].strip():
                has_content[i] = True

    # Always keep the first few critical columns and header
    return has_content


def filter_columns(header, rows):
    """Filter out completely empty columns but always keep headers"""
    has_content = get_non_empty_columns(header, rows)

    # Keep columns that have content
    filtered_header = [header[i] for i in range(len(header)) if has_content[i]]
    filtered_rows = []

    last_was_empty = False
    for row in rows:
        if is_empty_row(row):
            # Only add one empty row if there are consecutive empty rows
            # and only if we already have some content (not at the beginning)
            if not last_was_empty and filtered_rows:
                filtered_rows.append([""] * len(filtered_header))
                last_was_empty = True
        else:
            filtered_row = []
            for i in range(len(header)):
                if has_content[i]:
                    if i < len(row):
                        filtered_row.append(row[i])
                    else:
                        filtered_row.append("")
            filtered_rows.append(filtered_row)
            last_was_empty = False

    # Remove trailing empty row if present
    if filtered_rows and is_empty_row(filtered_rows[-1]):
        filtered_rows.pop()

    return filtered_header, filtered_rows


def organize_work_by_person(rows):
    """Organize work items by person responsible"""
    if not rows:
        return {}

    # First row is the header
    header = rows[0]
    data_rows = rows[1:]

    # Find the "FOLLOW UP" column index
    follow_up_index = None
    for i, col in enumerate(header):
        if "FOLLOW UP" in col.upper():
            follow_up_index = i
            break

    if follow_up_index is None:
        raise ValueError("Could not find 'FOLLOW UP' column in CSV")

    # Dictionary to store work items for each person
    person_work = {}

    for row in data_rows:
        # Check if this is an empty row (divider)
        if is_empty_row(row):
            # Add empty row to all people's lists
            for person in person_work:
                person_work[person].append(row)
            continue

        # Get the names from the follow-up column
        if len(row) > follow_up_index:
            names_string = row[follow_up_index]
            names = parse_names(names_string)

            # Add this row to each person's list
            for name in names:
                if name not in person_work:
                    person_work[name] = []
                person_work[name].append(row)

    return header, person_work


def create_pdf_for_person(person_name, header, rows, output_path):
    """Create a PDF with work items for a specific person"""
    # Filter out empty columns first
    filtered_header, filtered_rows = filter_columns(header, rows)

    # Create the PDF document
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        rightMargin=0.4 * inch,
        leftMargin=0.4 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.4 * inch,
    )

    # Container for the 'Flowable' objects
    elements = []

    # Define styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Heading1"],
        fontSize=14,
        textColor=colors.HexColor("#1a1a1a"),
        spaceAfter=12,
        alignment=1,  # Center alignment
    )

    # Style for table cells with wrapping
    cell_style = ParagraphStyle(
        "CellStyle",
        parent=styles["Normal"],
        fontSize=6,
        leading=7.5,
        wordWrap="CJK",
    )

    header_cell_style = ParagraphStyle(
        "HeaderCellStyle",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        textColor=colors.whitesmoke,
        fontName="Helvetica-Bold",
        alignment=1,  # Center
    )

    # Add title
    title = Paragraph(f"Work Allocation - {person_name}", title_style)
    elements.append(title)
    elements.append(Spacer(1, 0.1 * inch))

    # Prepare data for the table with Paragraph objects for wrapping
    table_data = []

    # Add header row with Paragraphs
    header_row = [Paragraph(str(cell), header_cell_style) for cell in filtered_header]
    table_data.append(header_row)

    # Add data rows with Paragraphs
    for row in filtered_rows:
        if is_empty_row(row):
            # Empty divider row
            table_data.append([""] * len(filtered_header))
        else:
            # Wrap each cell in a Paragraph for text wrapping
            formatted_row = []
            for cell in row:
                cell_text = str(cell).strip()
                if cell_text:
                    formatted_row.append(Paragraph(cell_text, cell_style))
                else:
                    formatted_row.append("")
            table_data.append(formatted_row)

    # Create the table
    # Calculate column widths based on content
    page_width = landscape(A4)[0] - 0.8 * inch
    num_cols = len(filtered_header)

    # Create a mapping of column names to their preferred widths
    col_width_map = {
        "Working Date": 0.65 * inch,
        "DATE": 0.65 * inch,
        "TIME": 0.5 * inch,
        "PROGRAM": 0.8 * inch,
        "THINGS TO DO": 1.8 * inch,
        "PAYMENT": 0.6 * inch,
        "CONTACT": 0.7 * inch,
        "FOLLOW UP": 1.0 * inch,
        "Work Head": 0.65 * inch,
        "LIST": 0.45 * inch,
    }

    # Calculate column widths
    col_widths = []
    fixed_width_total = 0
    flexible_cols = 0

    for col_name in filtered_header:
        if col_name in col_width_map:
            col_widths.append(col_width_map[col_name])
            fixed_width_total += col_width_map[col_name]
        else:
            col_widths.append(None)  # Will be calculated later
            flexible_cols += 1

    # Distribute remaining width to flexible columns
    if flexible_cols > 0:
        remaining_width = page_width - fixed_width_total
        flexible_width = remaining_width / flexible_cols
        col_widths = [w if w is not None else flexible_width for w in col_widths]
    else:
        # If all columns have fixed widths but don't fill the page, scale them
        if fixed_width_total < page_width:
            scale = page_width / fixed_width_total
            col_widths = [w * scale for w in col_widths]

    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    # Add style to the table
    table_style = TableStyle(
        [
            # Header style
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
            ("TOPPADDING", (0, 0), (-1, 0), 8),
            # Data rows style
            ("BACKGROUND", (0, 1), (-1, -1), colors.white),
            ("ALIGN", (0, 1), (-1, -1), "LEFT"),
            ("TOPPADDING", (0, 1), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
            ("LEFTPADDING", (0, 1), (-1, -1), 4),
            ("RIGHTPADDING", (0, 1), (-1, -1), 4),
            # Grid
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            # Alternating row colors for better readability
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [colors.white, colors.HexColor("#F2F2F2")],
            ),
        ]
    )

    # Add special styling for empty rows (dividers)
    for i, row in enumerate(filtered_rows, start=1):  # Start at 1 to skip header
        if is_empty_row(row):
            table_style.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#D9D9D9"))
            table_style.add("LINEABOVE", (0, i), (-1, i), 2, colors.HexColor("#808080"))
            table_style.add("LINEBELOW", (0, i), (-1, i), 2, colors.HexColor("#808080"))

    table.setStyle(table_style)
    elements.append(table)

    # Build the PDF
    doc.build(elements)
    print(f"Generated PDF for {person_name}: {output_path}")


def main():
    # Define paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(script_dir, "work.csv")
    output_dir = os.path.join(script_dir, "work")

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Read and organize the work data
    print("Reading work.csv...")
    rows = read_work_csv(csv_path)

    print("Organizing work by person...")
    header, person_work = organize_work_by_person(rows)

    # Generate PDF for each person
    print(f"\nGenerating PDFs for {len(person_work)} people...\n")

    for person_name, work_rows in sorted(person_work.items()):
        # Create a safe filename
        safe_filename = person_name.replace("/", "-").replace("\\", "-")
        output_path = os.path.join(output_dir, f"{safe_filename}.pdf")

        # Generate the PDF
        create_pdf_for_person(person_name, header, work_rows, output_path)

    print(f"\n✓ Successfully generated {len(person_work)} PDFs in {output_dir}")


if __name__ == "__main__":
    main()
