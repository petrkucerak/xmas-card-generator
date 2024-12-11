# Christmas Card Generator Script

## Overview

This Python script generates personalized Christmas cards in a Word document format (`.docx`) using data from a CSV file. It provides tailored messages for two types of recipients (`A` and `B`) based on the input data and applies custom formatting to produce professional-looking cards.

---

## Features

- **Customizable Messages**: Generates cards with specific content for `type A` (informal tone) and `type B` (formal tone) recipients.
- **Personalized Quotes**: Includes unique quotes from the CSV data, formatted and centered on the card.
- **A4 Page Layout**: Configures the Word document to use A4 paper size with custom margins.
- **Styled Text**: Sets default font style, size, and line spacing for professional formatting.
- **Multi-page Output**: Each card is created on a separate page in the output document.

---

## Prerequisites

1. **Python 3.x**: Install the latest version of Python.
2. **Required Libraries**: Install the following Python packages:
   - `python-docx`

   You can install the dependencies using pip:
   ```bash
   pip install python-docx
   ```

3. **Input CSV File**: The script reads data from a CSV file (`data.csv`) with the following columns:
   - `type`: Recipient type (`A` or `B`)
   - `address`: Personalized address or salutation
   - `quote`: The unique quote for the recipient
   - `source`: Source/reference for the quote

---

## Usage

1. **Prepare the CSV File**: Create a `data.csv` file in the same directory as the script with the required columns.

   Example:
   ```csv
   type,address,quote,source
   A,"John Doe","For I know the plans I have for you.","Jeremiah 29:11"
   B,"Jane Smith","The Lord is my shepherd; I shall not want.","Psalm 23:1"
   ```

2. **Run the Script**: Execute the script to generate the Christmas cards.
   ```bash
   python script.py
   ```

3. **Output**: The script generates a Word document named `xmas-cards.docx` in the same directory.

---

## Customization

### Static Messages
The script includes predefined messages for `type A` and `type B` recipients. To update these messages:
- Modify the `head_a`, `body_a`, `name_a` variables for type `A`.
- Modify the `head_b`, `body_b`, `name_b` variables for type `B`.

### Formatting
- Adjust the default font and line spacing by editing the `style` settings in the script.

---

## Error Handling

- If the input CSV file is not found, the script outputs an error message: `Error: File not found at {file_path}`.
- General exceptions are caught and displayed for troubleshooting.

---

## Output Example

For an entry in the CSV file:
```csv
type,address,quote,source
A,"Dear John","You are the light of the world.","Matthew 5:14"
```

The script generates a card with:
- A predefined message for `type A`.
- Personalized salutation: `Dear John`.
- Quote: `“You are the light of the world.”`
- Quote reference: `Matthew 5:14`.

---

## License

This script is open-source and available for personal or community use. Feel free to adapt and share as needed.