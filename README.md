# BulkEdit - PDF Name Personalization

This tool takes a PDF template and creates personalized copies by adding names to the first page.

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Install Poppler (required for pdf2image):
   - **macOS**: `brew install poppler`
   - **Ubuntu/Debian**: `sudo apt-get install poppler-utils`
   - **Windows**: Download from [poppler releases](https://github.com/oschwartz10612/poppler-windows/releases/)

## Usage

1. Place your PDF template files in the `base_cards/` directory
2. Create a `names.csv` file with 4 columns:
   - **Column 1**: Subfolder name (inside invites/)
   - **Column 2**: (ignored - can be anything)
   - **Column 3**: Name to print on card
   - **Column 4**: Base card filename (from base_cards/)
   
   Example:
   ```csv
   wedding,ignored,John Smith,wedding_card.pdf
   pooja,ignored,Jane Doe,pooja_card.pdf
   reception,ignored,Bob Johnson,reception_card.pdf
   ```

3. Adjust configuration in `main.py`:
   - `BASE_CARDS_DIR`: Directory containing base PDF cards
   - `TEXT_Y_POSITION`: Y coordinate where names should appear
   - `FONT_SIZE`: Size of the text
   - `TEXT_COLOR`: Color of the text

4. Run the script:
```bash
python main.py
```

5. Find your personalized PDFs in the `invites/` directory

## Configuration Options

- `BASE_CARDS_DIR`: Directory containing base PDF card templates (default: "base_cards")
- `NAMES_CSV`: Path to CSV file with 4 columns: folder, skip, name, card
- `OUTPUT_DIR`: Directory for output PDFs (default: "invites")
- `FONT_SIZE`: Font size for names (default: 60)
- `TEXT_Y_POSITION`: Y coordinate for text placement (default: 1000)
- `TEXT_COLOR`: Color of the text (e.g., "black", "#722F37" for wine)
- `CENTER_TEXT`: Whether to center text horizontally (default: True)
- `MAX_TEXT_WIDTH`: Maximum width before splitting into 2 lines (default: 800px)
- `LINE_SPACING`: Spacing between lines when split (default: 100px)

