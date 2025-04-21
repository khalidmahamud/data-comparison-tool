# Data Comparison Tool

A powerful web-based application for comparing and analyzing textual data in Excel files, with support for intelligent AI-assisted corrections and contextual analysis.

## Features

- **Interactive Comparison Interface**: Side-by-side visualization of primary and secondary text columns with difference highlighting
- **AI-Powered Text Analysis**: Utilizes Google's Gemini models to analyze and suggest improvements to content
- **Excel File Handling**:
  - Split large Excel files into manageable chunks
  - Merge processed chunks back into a single file
  - Preserve all Excel styles during processing
- **Visual Difference Highlighting**: See exactly what changed between text versions
- **Similarity Ratio Calculation**: Automatic calculation of similarity between text columns
- **Commenting System**: Add and save comments on individual entries
- **Arabic Text Support**: Special handling for Arabic text content
- **Flexible Configuration**: YAML-based configuration for easy customization

## Getting Started

### Prerequisites

- Python 3.7+
- Flask
- Pandas
- OpenPyXL
- PyYAML

### Installation

1. Clone the repository:

```
git clone https://github.com/your-username/data-comparison-tool.git
cd data-comparison-tool
```

2. Install required packages:

```
pip install -r requirements.txt
```

3. Configure your settings in `config_flash.yaml`

### Usage

1. **Start the web application**:

```
python app.py
```

2. **Split large Excel files into chunks**:

```
python sm.py
```

3. Open your browser and navigate to `http://localhost:5000`

## Configuration

The application is configured through `config_flash.yaml`:

```yaml
processing:
  batch_size: 5
  max_retries: 3
  retry_delay: 0
  save_interval: 5
  start_row: 0

api_settings:
  gemini_flash_1:
    api_key: "YOUR_API_KEY_HERE"
    model: "gemini-2.0-flash"
    max_tokens: 8192

file_settings:
  input_file: "input.xlsx"
  output_file: "output.xlsx"
  prompts_file: "prompts.txt"
  chunks_directory: "chunks"
  merged_file: "merged_output.xlsx"
  rows_per_chunk: 500
  action: "merge" # Options: "split" or "merge"

excel_settings:
  sheet_name: "data"
  columns:
    primary_text: "hadith_details"
    secondary_text: "analysis-3"
    ratio: "ratio"
    number: "number"
    arabic_text: "arabic_text"
```

### Configuration Details

#### Processing

- **batch_size**: Number of rows to process in a single batch for AI operations
- **max_retries**: Maximum number of retry attempts if an API call fails
- **retry_delay**: Delay in seconds between retry attempts
- **save_interval**: Number of processed rows after which to save progress
- **start_row**: Row number to start processing from (useful for resuming)

#### API Settings

- **gemini_flash_1**: Configuration for the primary Gemini model
  - **api_key**: Your Google Gemini API key
  - **model**: The specific Gemini model to use (e.g., "gemini-2.0-flash")
  - **max_tokens**: Maximum token limit for generated responses

You can configure multiple model entries (gemini_flash_2, gemini_flash_3, etc.) for load balancing.

#### File Settings

- **input_file**: Path to the source Excel file to process
- **output_file**: Path where processed data will be saved
- **prompts_file**: Path to file containing AI prompts
- **chunks_directory**: Directory where file chunks will be stored
- **merged_file**: Name of the final merged output file
- **rows_per_chunk**: Maximum number of rows in each Excel chunk
- **action**: Operation to perform: "split" (divide file into chunks) or "merge" (combine chunks)

#### Excel Settings

- **sheet_name**: Name of the worksheet to process
- **columns**: Mapping of logical column names to actual Excel column headers
  - **primary_text**: Column containing the primary text data
  - **secondary_text**: Column containing the comparison text data
  - **ratio**: Column where similarity ratios will be stored
  - **number**: Column containing row/entry numbers or IDs
  - **arabic_text**: Column containing Arabic text (if applicable)

## Project Structure

- `app.py`: Main Flask web application
- `sm.py`: Excel file splitting and merging utility
- `src/`: Source code directory
  - `generate_cell.py`: AI text generation utilities
  - `config.py`: Configuration management
  - `ai.py`: AI integration module
  - `prompt.py`: Prompt handling utilities
- `templates/`: HTML templates for the web interface
- `static/`: Static assets (CSS, JavaScript, images)
- `chunks/`: Directory for storing Excel file chunks
- `prompts/`: Directory for prompt templates

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
