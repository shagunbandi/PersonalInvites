import csv
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
    HRFlowable,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO


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

    # Skip specific columns: 1st (index 0), 6th (index 5), 10th (index 9)
    columns_to_skip = {0, 5, 9}

    # Keep columns that have content and are not in the skip list
    filtered_header = [
        header[i]
        for i in range(len(header))
        if has_content[i] and i not in columns_to_skip
    ]
    filtered_rows = []

    for row in rows:
        if is_empty_row(row):
            # Skip all empty rows
            continue
        else:
            filtered_row = []
            for i in range(len(header)):
                if has_content[i] and i not in columns_to_skip:
                    if i < len(row):
                        filtered_row.append(row[i])
                    else:
                        filtered_row.append("")
            filtered_rows.append(filtered_row)

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


class WatermarkCanvas(canvas.Canvas):
    """Custom canvas that adds a watermark to each page"""

    def __init__(self, *args, **kwargs):
        self.logo_path = kwargs.pop("logo_path", None)
        self._watermark_image = None
        canvas.Canvas.__init__(self, *args, **kwargs)
        # Preload the watermark image
        if self.logo_path and os.path.exists(self.logo_path):
            self._load_watermark()

    def _load_watermark(self):
        """Load and prepare the watermark image once"""
        try:
            from PIL import Image

            # Check if file is PDF or image
            if self.logo_path.lower().endswith(".pdf"):
                from pdf2image import convert_from_path

                # Convert first page of PDF to image
                images = convert_from_path(
                    self.logo_path, first_page=1, last_page=1, dpi=150
                )
                if images:
                    # Save to BytesIO to use with ImageReader
                    img_buffer = BytesIO()
                    images[0].save(img_buffer, format="PNG")
                    img_buffer.seek(0)

                    self._watermark_image = ImageReader(img_buffer)
                    self._watermark_size = images[0].size
            else:
                # Load image directly (PNG, JPG, etc.)
                img = Image.open(self.logo_path)
                self._watermark_image = ImageReader(self.logo_path)
                self._watermark_size = img.size
        except Exception as e:
            print(f"Warning: Could not load watermark: {e}")
            self._watermark_image = None

    def showPage(self):
        """Override showPage to finalize the page"""
        canvas.Canvas.showPage(self)

    def _drawBeforeContent(self):
        """Draw the watermark BEFORE any content"""
        if self._watermark_image is None:
            return

        self.saveState()

        # Set transparency (0.0 = transparent, 1.0 = opaque)
        self.setFillAlpha(0.6)  # 60% opacity for better visibility
        self.setStrokeAlpha(0.6)

        # Get page dimensions
        page_width, page_height = A4

        try:
            img_width, img_height = self._watermark_size

            # Scale logo to fit centered on page (use 60% of page size)
            scale = min(
                (page_width * 0.6) / img_width, (page_height * 0.6) / img_height
            )
            scaled_width = img_width * scale
            scaled_height = img_height * scale

            # Center the logo on the page
            x = (page_width - scaled_width) / 2
            y = (page_height - scaled_height) / 2

            # Draw the logo
            self.drawImage(
                self._watermark_image,
                x,
                y,
                width=scaled_width,
                height=scaled_height,
                mask="auto",
                preserveAspectRatio=True,
            )
        except Exception as e:
            print(f"Warning: Could not draw watermark in old canvas: {e}")

        self.restoreState()

    def _startPage(self):
        """Override _startPage to draw watermark first"""
        canvas.Canvas._startPage(self)
        self._drawBeforeContent()


def add_watermark_to_page(canvas_obj, doc, logo_path):
    """Add watermark to a page using canvas callback"""
    try:
        from PIL import Image

        # Load the logo
        img = Image.open(logo_path)
        img_reader = ImageReader(logo_path)
        img_width, img_height = img.size

        # Get page dimensions
        page_width, page_height = A4

        # Scale logo to fit centered on page (use 60% of page size)
        scale = min((page_width * 0.6) / img_width, (page_height * 0.6) / img_height)
        scaled_width = img_width * scale
        scaled_height = img_height * scale

        # Center the logo on the page
        x = (page_width - scaled_width) / 2
        y = (page_height - scaled_height) / 2

        # Save state and set transparency
        canvas_obj.saveState()
        canvas_obj.setFillAlpha(0.1)  # 50% opacity so watermark shows on top
        canvas_obj.setStrokeAlpha(0.1)

        # Draw the logo
        canvas_obj.drawImage(
            img_reader,
            x,
            y,
            width=scaled_width,
            height=scaled_height,
            mask="auto",
            preserveAspectRatio=True,
        )

        canvas_obj.restoreState()
    except Exception as e:
        print(f"Error adding watermark: {e}")


def create_pdf_for_person(person_name, header, rows, output_path, logo_path=None):
    """Create a PDF with work items for a specific person"""
    # Filter out empty columns first
    filtered_header, filtered_rows = filter_columns(header, rows)

    # Create the PDF document
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=0.4 * inch,
        leftMargin=0.4 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.4 * inch,
    )

    # Set up watermark callback if logo exists
    watermark_func = None
    if logo_path and os.path.exists(logo_path):
        watermark_func = lambda c, d: add_watermark_to_page(c, d, logo_path)

    # Container for the 'Flowable' objects
    elements = []

    # Define styles
    styles = getSampleStyleSheet()

    # Main title style with decorative line
    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Heading1"],
        fontSize=18,
        textColor=colors.HexColor("#2C3E50"),
        spaceAfter=8,
        spaceBefore=4,
        alignment=1,  # Center alignment
        fontName="Helvetica-Bold",
    )

    # Subtitle style for "Work Allocation"
    subtitle_style = ParagraphStyle(
        "SubtitleStyle",
        parent=styles["Normal"],
        fontSize=10,
        textColor=colors.HexColor("#7F8C8D"),
        spaceAfter=12,
        alignment=1,  # Center alignment
        fontName="Helvetica",
        letterSpacing=2,
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

    # Add beautiful title with person's name
    elements.append(Spacer(1, 0.05 * inch))
    subtitle = Paragraph("WORK ALLOCATION", subtitle_style)
    elements.append(subtitle)

    # Add decorative line
    line = HRFlowable(
        width="30%",
        thickness=1,
        color=colors.HexColor("#3498DB"),
        spaceAfter=8,
        spaceBefore=2,
        hAlign="CENTER",
    )
    elements.append(line)

    # Add person's name as main title
    name_title = Paragraph(person_name, title_style)
    elements.append(name_title)

    # Add another decorative line
    line2 = HRFlowable(
        width="30%",
        thickness=1,
        color=colors.HexColor("#3498DB"),
        spaceAfter=12,
        spaceBefore=8,
        hAlign="CENTER",
    )
    elements.append(line2)

    # Prepare data for the table with Paragraph objects for wrapping
    table_data = []

    # Add header row with Paragraphs
    header_row = [Paragraph(str(cell), header_cell_style) for cell in filtered_header]
    table_data.append(header_row)

    # Add data rows with Paragraphs
    for row in filtered_rows:
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
    page_width = A4[0] - 0.8 * inch
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
        "FOLLOW UP": 1.8 * inch,  # Increased width for multiple names
        "Work Head": 0.45 * inch,  # Reduced width for short labels
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

    # Add style to the table with transparent backgrounds so watermark shows through
    table_style = TableStyle(
        [
            # Header style
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
            ("TOPPADDING", (0, 0), (-1, 0), 8),
            # Data rows style - NO BACKGROUND so watermark shows through
            ("ALIGN", (0, 1), (-1, -1), "LEFT"),
            ("TOPPADDING", (0, 1), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
            ("LEFTPADDING", (0, 1), (-1, -1), 4),
            ("RIGHTPADDING", (0, 1), (-1, -1), 4),
            # Grid
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
    )

    table.setStyle(table_style)
    elements.append(table)

    # Build the PDF with watermark callback
    if watermark_func:
        doc.build(elements, onFirstPage=watermark_func, onLaterPages=watermark_func)
    else:
        doc.build(elements)
    print(f"Generated PDF for {person_name}: {output_path}")


def main():
    # Define paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(script_dir, "work.csv")
    output_dir = os.path.join(script_dir, "work")
    logo_path = os.path.join(script_dir, "logo.png")

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Check if logo exists
    if os.path.exists(logo_path):
        print(f"Found logo at: {logo_path}")
    else:
        print("Warning: logo.pdf not found, PDFs will be generated without watermark")

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

        # Generate the PDF with watermark
        create_pdf_for_person(person_name, header, work_rows, output_path, logo_path)

    print(f"\n✓ Successfully generated {len(person_work)} PDFs in {output_dir}")


if __name__ == "__main__":
    main()
