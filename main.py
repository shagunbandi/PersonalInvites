from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path
import csv
import os

# Configuration
BASE_CARDS_DIR = "base_cards"  # Directory containing base PDF cards
NAMES_CSV = "names.csv"
OUTPUT_DIR = "invites"
FONT_SIZE = 60
TEXT_Y_POSITION = 1000  # Y coordinate for the name (distance from top)
TEXT_X_POSITION = 500  # X coordinate (only used if CENTER_TEXT is False)
TEXT_COLOR = "#722F37"  # Wine color
CENTER_TEXT = True  # Set to True to center text horizontally
MAX_TEXT_WIDTH = 800  # Maximum width before splitting into 2 lines (in pixels)
LINE_SPACING = 100  # Spacing between lines when split

# Font options (tries Calibri first, then falls back to system fonts)
FONT_OPTIONS = [
    "fonts/calisto-mt.ttf",  # Place Calibri.ttf in project directory
    "/Library/Fonts/Calibri.ttf",  # Microsoft Office font
    "/System/Library/Fonts/Supplemental/Calibri.ttf",
    "~/Library/Fonts/Calibri.ttf",  # User fonts
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/System/Library/Fonts/Helvetica.ttc",
    "/System/Library/Fonts/SFNSDisplay.ttf",
    "/Library/Fonts/Arial.ttf",
]

# Create output directory if it doesn't exist
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Load font - try multiple options
font = None
for font_path in FONT_OPTIONS:
    try:
        # Expand user path (~ to home directory)
        expanded_path = os.path.expanduser(font_path)
        font = ImageFont.truetype(expanded_path, FONT_SIZE)
        print(f"Using font: {font_path}")
        break
    except OSError:
        continue

if font is None:
    print("Warning: Could not load TrueType font, using default font")
    font = ImageFont.load_default()

# Cache for loaded PDFs to avoid re-converting the same card multiple times
pdf_cache = {}


def load_pdf_card(card_name):
    """Load a PDF card from the base_cards directory, using cache if available"""
    if card_name in pdf_cache:
        return pdf_cache[card_name]

    card_path = os.path.join(BASE_CARDS_DIR, card_name)
    if not os.path.exists(card_path):
        # Try adding .pdf extension if not present
        if not card_path.endswith(".pdf"):
            card_path += ".pdf"

    if not os.path.exists(card_path):
        raise FileNotFoundError(f"Base card not found: {card_name}")

    print(f"  Loading base card: {card_name}")
    pdf_pages = convert_from_path(card_path, dpi=200)
    pdf_cache[card_name] = pdf_pages
    return pdf_pages


# Process each name
print(f"Processing names from: {NAMES_CSV}")
with open(NAMES_CSV) as f:
    reader = csv.reader(f)
    processed_count = 0

    for row_num, row in enumerate(reader, 1):
        # Skip empty lines
        if not row or len(row) < 4:
            continue

        # Read CSV columns: folder, skip, name, card
        folder_name = row[0].strip()
        # row[1] is skipped
        name = row[2].strip()
        card_name = row[3].strip()

        if not name or not card_name:
            print(f"⚠ Skipping row {row_num}: Missing name or card")
            continue

        print(f"Processing: {name} → {card_name} (folder: {folder_name})")

        # Load the appropriate base card
        try:
            pdf_pages = load_pdf_card(card_name)
        except FileNotFoundError as e:
            print(f"  ✗ Error: {e}")
            continue

        # Create a copy of the first page and add the name
        first_page = pdf_pages[0].copy()
        draw = ImageDraw.Draw(first_page)

        # Check if text needs to be split into multiple lines
        bbox = draw.textbbox((0, 0), name, font=font)
        text_width = bbox[2] - bbox[0]

        if text_width > MAX_TEXT_WIDTH:
            # Split name into 2 lines - try to split at a space near the middle
            words = name.split()
            if len(words) > 1:
                # Find the best split point (closest to middle)
                mid_point = len(name) // 2
                best_split = 1
                min_diff = float("inf")

                for i in range(1, len(words)):
                    split_pos = len(" ".join(words[:i]))
                    diff = abs(split_pos - mid_point)
                    if diff < min_diff:
                        min_diff = diff
                        best_split = i

                line1 = " ".join(words[:best_split])
                line2 = " ".join(words[best_split:])
                lines = [line1, line2]
            else:
                # Single long word, just use it as is
                lines = [name]
        else:
            lines = [name]

        # Draw each line centered
        page_width = first_page.width
        start_y = TEXT_Y_POSITION - ((len(lines) - 1) * LINE_SPACING / 2)

        for i, line in enumerate(lines):
            if CENTER_TEXT:
                # Get the bounding box of the line
                bbox = draw.textbbox((0, 0), line, font=font)
                line_width = bbox[2] - bbox[0]
                x_position = (page_width - line_width) / 2
                y_position = start_y + (i * LINE_SPACING)
            else:
                x_position = TEXT_X_POSITION
                y_position = start_y + (i * LINE_SPACING)

            draw.text((x_position, y_position), line, font=font, fill=TEXT_COLOR)

        # Create subfolder if specified
        if folder_name:
            output_folder = os.path.join(OUTPUT_DIR, folder_name)
            os.makedirs(output_folder, exist_ok=True)
        else:
            output_folder = OUTPUT_DIR

        # Save all pages as a PDF
        output_path = os.path.join(output_folder, f"{name}.pdf")

        # Combine first page with remaining pages
        all_pages = [first_page] + pdf_pages[1:]

        # Save as PDF
        all_pages[0].save(
            output_path,
            save_all=True,
            append_images=all_pages[1:],
            resolution=200.0,
            quality=95,
        )

        print(f"  ✓ Created: {output_path}")
        processed_count += 1

print(
    f"\n✅ Done! Created {processed_count} personalized PDFs in '{OUTPUT_DIR}/' directory"
)
