# Excel Column Comparison Tool

This Flask application compares data between columns A and B in an Excel file and displays the differences in a web interface.

## Features

- Reads data from an Excel file named `input.xlsx`
- Compares values between columns A and B
- Highlights differences with color coding (red for different, green for same)
- Allows selecting the number of rows to display (10, 50, or 100)
- Includes pagination for navigating through large datasets
- **NEW**: API-based regeneration of column B values using Gemini AI

## Installation

1. Make sure you have Python 3.7+ installed
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Place your Excel file named `input.xlsx` in the same directory as the application
2. The Excel file should have columns labeled 'hadith_details' and 'analysis-3' in the 'hadith' sheet
3. Create a `config_flash.yaml` file for the API configuration (see Configuration section)
4. Create a `prompts.txt` file with your prompts for text generation
5. Run the application:

```bash
python app.py
```

6. Open your web browser and navigate to `http://127.0.0.1:5000/`
7. Use the dropdown to select how many rows to display per page
8. Navigate through pages using the pagination controls at the bottom
9. Regenerate column B content using the regenerate button (autorenew icon)

## Configuration

Create a `config_flash.yaml` file with the following structure:

```yaml
processing:
  batch_size: 5
  max_retries: 3
  retry_delay: 0
  start_row: 0

api_settings:
  gemini_flash_1:
    api_key: "YOUR_GEMINI_API_KEY"
    model: "gemini-2.0-flash"
    max_tokens: 8192

file_settings:
  input_file: "input.xlsx"
  output_file: "input.xlsx"
  prompts_file: "prompts.txt"
  chunks_directory: "chunks"
  merged_file: "merged_output.xlsx"

excel_settings:
  sheet_name: "hadith"
  columns:
    primary_text: "hadith_details"
    secondary_text: "analysis-3"
    ratio: "ratio"
    number: "number"
    arabic_text: "arabic_text"
```

The `excel_settings` section allows you to customize:

- `sheet_name`: The name of the Excel sheet containing your data
- `columns`: Mapping of column logical names to actual column names in Excel
  - `primary_text`: Column containing the primary text (default: "hadith_details")
  - `secondary_text`: Column containing the secondary text for comparison (default: "analysis-3")
  - `ratio`: Column for storing similarity ratios (default: "ratio")
  - `number`: Column containing row identifiers (default: "number")
  - `arabic_text`: Column containing Arabic text if available (default: "arabic_text")

## Creating a Sample Excel File

You can create a sample Excel file for testing using the provided script:

```bash
python create_sample_data.py
```

Or you can manually create an Excel file with the following structure:

- Create a file named `input.xlsx`
- Add a sheet named 'hadith'
- Add columns with headers 'hadith_details' and 'analysis-3'
- Add data to these columns for comparison

## Example

The application will display a table with three columns:

- Row number
- Column A value (hadith_details)
- Column B value (analysis-3)

Cells will be highlighted in:

- Green when values in columns A and B are the same
- Red when values are different

## Cell Regeneration

To regenerate content in Column B:

1. Click the regenerate button (autorenew icon) in the cell you want to update
2. The application will call the Gemini API using your configured credentials
3. The new content will replace the existing content in Column B
4. Differences will be highlighted automatically
